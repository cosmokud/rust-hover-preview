//! The scheduled task that starts this app **elevated** at a logon.
//!
//! `Software\Microsoft\Windows\CurrentVersion\Run` is the ordinary way this app starts with
//! Windows, and it is also the one way it cannot start elevated: Windows runs a `Run` value with
//! the logged-on token and never raises a UAC prompt for one, so an entry whose executable asks
//! for administrator rights starts **nothing at all**. A logon that starts this app as
//! administrator therefore needs a mechanism that can carry a run level, and there is exactly one
//! of those in the Shell's own vocabulary: a Task Scheduler task with a logon trigger and
//! `/RL HIGHEST` — "Run with highest privileges".
//!
//! **One task, one name, one folder.** `\RustHoverPreview` is a fact of the machine the way the
//! `Run` value's name is: the row of the tray, the removal, and anyone reading Task Scheduler all
//! know this app by it. It is machine-wide, so a second account that turns the row on replaces it
//! rather than adding beside it, and the reconciler in `app::startup` rewrites it whenever the
//! copy of this app that is running is not the one it names.
//!
//! **Registering is a machine write and needs the elevated token.** Measured on Windows 11:
//! `schtasks /Create` and `TaskService::RegisterTaskDefinition` are both refused with
//! `E_ACCESSDENIED` on a task folder in the root for a `HighestAvailable` principal, and
//! `LeastPrivilege` in the root is refused too — the root folder's ACL is what decides it, not the
//! run level. So every write here answers [`Refused::NeedsElevation`] where this process is not
//! elevated, and what raises the prompt is the caller (`app::startup::raise_to_elevation`).
//!
//! **And the token is not the only thing a write needs: COM has to be up.** `CoCreateInstance`
//! answers `CO_E_NOTINITIALIZED` on a thread that never joined an apartment, and two of the threads
//! that write a task have not — the elevated helper does its one write and exits before the tray's
//! `CoInitializeEx`, and the reconciler at a start runs ahead of it. That is why the apartment is
//! made here rather than asked of the caller (see [`connect`]).
//!
//! **Reading needs nothing.** A task's XML is exported by the Task Scheduler service to any caller
//! that may read the folder, and the service keeps its own copy under
//! `%SystemRoot%\System32\Tasks\<name>` — a UTF-16 file this app can read unelevated. That is what
//! lets the tray draw a checkmark from the machine on every opening of its menu without a prompt,
//! and the file is the fallback for the day the service refuses to talk (see [`facts`]).
//!
//! **The XML is written here rather than by the COM builder.** A definition assembled through
//! `ITaskDefinition` is eight interfaces and a dozen property setters, and every one of them is a
//! place this app could set a value the platform then reinterprets; the text below is the same
//! document `schtasks /XML` and `Export-ScheduledTask` produce, in one place, held to a test.
//! What the task says, in words:
//!
//! - **the trigger is a logon of one account** — the account whose desktop this is
//!   (`app::elevation::desktop_user_name`), so a machine with two accounts starts one copy per
//!   account that asked for it rather than one per logon of anybody;
//! - **the principal runs at `HighestAvailable`** with the interactive token, which is the whole
//!   point: the split token of an administrator's own logon, raised, with no prompt;
//! - **`IgnoreNew`**, so a copy already standing is not started twice;
//! - **`ExecutionTimeLimit` of `PT0S`**, which is "no limit": a task whose process is a tray app
//!   is one the Task Scheduler would otherwise end after three days;
//! - **`StopIfGoingOnBatteries` and `DisallowStartIfOnBatteries` both off**, because a laptop at
//!   a logon is on a battery more often than not;
//! - and **no restart-on-failure**, because the app's own start is not something a task can fix
//!   and a restart loop over a broken install is worse than a task that failed.
//!
//! The action is the exe's own path, quoted by the XML's own encoding rather than by
//! `CommandLineToArgvW`'s rules: `Exec/Command` is a path, not a command line, so a folder with a
//! space in it needs no quotes here — and adding them would be a path with quotes in it.

use std::path::{Path, PathBuf};
use std::sync::OnceLock;

use windows::core::{BSTR, VARIANT};
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::{CoCreateInstance, CoIncrementMTAUsage, CLSCTX_INPROC_SERVER};
use windows::Win32::System::TaskScheduler::{
    ITaskService, TaskScheduler, TASK_CREATE_OR_UPDATE, TASK_LOGON_INTERACTIVE_TOKEN,
};

/// The task's name in the root folder, and the name every other reader of it knows this app by.
pub(crate) const TASK_NAME: &str = "RustHoverPreview";

/// What the task says of itself in Task Scheduler's own view, which is where a user who wonders
/// what this app put on their machine reads it.
const TASK_DESCRIPTION: &str =
    "Starts Rust Hover Preview when you sign in, with administrator rights \
     (the tray's System -> Run as administrator row).";

/// Why a write of the task did not happen.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Refused {
    /// The Task Scheduler refused this process's token, and nothing was written. It is the answer
    /// for an ordinary (unelevated) run on a machine where the task folder is the machine's,
    /// which is every machine: the write has to be made by an elevated process.
    NeedsElevation,
    /// The machine refused for a reason of its own — a policy, a broken Task Scheduler, a service
    /// that is not running. Nothing was written.
    NoWrite,
}

/// What the machine holds about the task, read without asking anybody for anything.
#[derive(Clone, PartialEq, Eq, Debug, Default)]
pub(crate) struct Facts {
    /// Whether a task of this name is there at all.
    pub present: bool,
    /// Whether a logon is allowed to run it — `Settings/Enabled` of the task, which Task Scheduler's
    /// own `Disable` and the Startup page's switch write. A task that is present and disabled
    /// starts nothing, exactly like a `Run` value Windows has been told not to run.
    pub enabled: bool,
    /// Whether it runs at the highest level available, which is the only setting here that makes a
    /// logon start this app as an administrator.
    pub elevated: bool,
    /// The path the task's action names, as the machine spells it.
    pub command: Option<String>,
}

impl Facts {
    /// Whether the machine would start **this** copy of the app, elevated, at a logon: the whole of
    /// what the tray's row is drawn from, and the whole of what a start checks before deciding
    /// there is nothing to write.
    pub(crate) fn starts_this_copy(&self, exe: &Path) -> bool {
        if !(self.present && self.enabled && self.elevated) {
            return false;
        }

        self.command
            .as_deref()
            .is_some_and(|command| same_path(command, exe))
    }
}

/// The task's own path in the Task Scheduler tree.
fn task_path() -> String {
    format!("\\{TASK_NAME}")
}

/// Where the Task Scheduler service keeps the task's XML, which is what a run may read without a
/// token of its own.
fn stored_file() -> Option<PathBuf> {
    let root = std::env::var_os("SystemRoot")?;
    Some(
        PathBuf::from(root)
            .join("System32")
            .join("Tasks")
            .join(TASK_NAME),
    )
}

/// What the machine holds about the task.
///
/// The service is asked first, because it is the truth about a task that was registered and then
/// edited from Task Scheduler's own window; the file is the fallback for a machine where the
/// service does not answer this process (a broken service, a policy), and it is stale in one known
/// way: the file is written when the task is registered and when the machine is told to change it,
/// and a change that the service applied but did not write back leaves an old document there — which
/// shows up as a path this app then rewrites, and never as a checkmark that lies about a task.
pub(crate) fn facts() -> Facts {
    read_through_the_service()
        .or_else(read_the_file)
        .unwrap_or_default()
}

/// Ask the Task Scheduler service for the task's XML.
fn read_through_the_service() -> Option<Facts> {
    let service = connect()?;

    // Safety: the folder path is this function's own string, and the interface it answers is what
    // the read below is made through.
    let folder = unsafe { service.GetFolder(&BSTR::from("\\")) }.ok()?;
    // Safety: as above, and a task that is not there answers an error rather than a null interface.
    let task = unsafe { folder.GetTask(&BSTR::from(TASK_NAME)) }.ok()?;

    // Safety: the interface above is a registered task, and its own XML is a string it hands over.
    let xml = unsafe { task.Xml() }.ok()?;
    // Safety: `Enabled` is a `VARIANT_BOOL` out-parameter of a live interface.
    let enabled = unsafe { task.Enabled() }.map(|flags| flags.as_bool());

    let facts = facts_of_xml(&xml.to_string())?;

    Some(Facts {
        present: true,
        enabled: enabled.unwrap_or(facts.enabled),
        ..facts
    })
}

/// Read the task out of the service's own copy of it.
fn read_the_file() -> Option<Facts> {
    let bytes = std::fs::read(stored_file()?).ok()?;
    let text = decode_utf16(&bytes)?;

    // The file is named for the task and is the service's own record of it, so its presence is the
    // task's — and its document says the rest. A file that was truncated, or one left half-written
    // by a machine that lost power, is not a document and is not a task either: an answer of
    // nothing is read as one the service is asked about rather than as a task of the wrong kind.
    facts_of_xml(&text)
}

/// The facts an exported task document states, or nothing where the text is not one.
///
/// It is the read the tray needs and the read a test can hold: the document is the platform's, so
/// what is taken from it is a path, two booleans and one string, each asked by its own element.
pub(crate) fn facts_of_xml(text: &str) -> Option<Facts> {
    let document = roxmltree::Document::parse(text).ok()?;
    let root = document.root_element();

    // A document that is not a task is not read as one with the defaults filled in: what this
    // answers decides whether the tray claims a machine will start this app at a logon.
    if !root.has_tag_name("Task") {
        return None;
    }

    let node = |name: &str| root.descendants().find(|node| node.has_tag_name(name));

    // **The two `Enabled` elements mean different things, and the task's own is the one that
    // matters.** `Settings/Enabled` is whether the task is enabled at all — Task Scheduler's own
    // `Disable` and the switch on **Settings → Apps → Startup** write it — while
    // `Triggers/LogonTrigger/Enabled` is whether that one trigger fires. A task whose trigger was
    // switched off starts nothing, so an answer of "there" has to hold both.
    let truthy = |node: Option<roxmltree::Node<'_, '_>>| {
        node.and_then(|element| element.text())
            .map(|text| text.trim().eq_ignore_ascii_case("true"))
    };

    let enabled = truthy(node("Settings").and_then(|settings| {
        settings
            .children()
            .find(|child| child.has_tag_name("Enabled"))
    }))
    // A task that says nothing about its own enable flag is one Task Scheduler's default applies
    // to, and the default is enabled.
    .unwrap_or(true)
        && truthy(node("LogonTrigger").and_then(|trigger| {
            trigger
                .children()
                .find(|child| child.has_tag_name("Enabled"))
        }))
        // And a trigger nobody wrote an enable flag for is a trigger that fires.
        .unwrap_or(true);

    let elevated = node("RunLevel")
        .and_then(|element| element.text())
        .is_some_and(|text| text.trim().eq_ignore_ascii_case("HighestAvailable"));

    // `Command` is the action's own path. The first one is the action this app wrote; a task
    // somebody edited to hold more than one is read by its first, which is what a logon runs first.
    let command = node("Command")
        .and_then(|element| element.text())
        .map(|text| text.trim().to_string())
        .filter(|text| !text.is_empty());

    Some(Facts {
        present: true,
        enabled,
        elevated,
        command,
    })
}

/// The text of a task file the Task Scheduler wrote: UTF-16, with or without its byte order mark.
fn decode_utf16(bytes: &[u8]) -> Option<String> {
    let body = bytes.strip_prefix(&[0xFF, 0xFE]).unwrap_or(bytes);
    let units: Vec<u16> = body
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();

    Some(String::from_utf16_lossy(&units))
}

/// Register the task, replacing one of the same name, and enable it.
///
/// The definition it writes is [`document`], and the principal it is handed is the desktop
/// account's — the one whose logon is to start this app — rather than this process's own, because an
/// elevated run answering another account's prompt is an account whose logon nobody sits at (see
/// `app::elevation`).
pub(crate) fn register(exe: &Path, user: &str) -> Result<(), Refused> {
    let service = connect().ok_or(Refused::NoWrite)?;

    // Safety: every interface here is a live member of the one above it, and each call is followed
    // by its own answer rather than assumed.
    let (folder, definition) = unsafe {
        let root = service
            .GetFolder(&BSTR::from("\\"))
            .map_err(|_| Refused::NoWrite)?;
        let definition = service.NewTask(0).map_err(|_| Refused::NoWrite)?;

        // The whole of the task is the document: `XmlText` is how a definition that came from a
        // document rather than from the setters beside it is put back into one.
        definition
            .SetXmlText(&BSTR::from(document(exe, user)))
            .map_err(|_| Refused::NoWrite)?;

        (root, definition)
    };

    // Safety: the definition was just written by this call and is not held by anything else.
    let registered = unsafe {
        folder.RegisterTaskDefinition(
            &BSTR::from(task_path()),
            &definition,
            TASK_CREATE_OR_UPDATE.0,
            &VARIANT::from(user),
            &VARIANT::default(),
            TASK_LOGON_INTERACTIVE_TOKEN,
            &VARIANT::default(),
        )
    };

    let task = registered.map_err(|error| {
        // `E_ACCESSDENIED` is the machine refusing this token rather than the write being wrong,
        // and it is the one answer a caller has something to do about.
        if error.code() == windows::Win32::Foundation::E_ACCESSDENIED {
            Refused::NeedsElevation
        } else {
            Refused::NoWrite
        }
    })?;

    // Safety: a registered task, and the flag is what Task Manager's own switch writes.
    unsafe { task.SetEnabled(VARIANT_BOOL(-1)) }.map_err(|_| Refused::NoWrite)?;

    Ok(())
}

/// Remove the task. A task that is not there is already what was asked for.
pub(crate) fn unregister() -> Result<(), Refused> {
    let service = connect().ok_or(Refused::NoWrite)?;

    // Safety: the folder is a live interface, and a delete of a name that is not there answers an
    // error this reads as settled — an absence is what a removal is for.
    let deleted = unsafe {
        let root = service
            .GetFolder(&BSTR::from("\\"))
            .map_err(|_| Refused::NoWrite)?;
        root.DeleteTask(&BSTR::from(TASK_NAME), 0)
    };

    match deleted {
        Ok(()) => Ok(()),
        Err(error) if error.code() == windows::Win32::Foundation::E_ACCESSDENIED => {
            Err(Refused::NeedsElevation)
        }
        // `HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)`: there is no such task, which is the state a
        // successful removal leaves behind.
        Err(error)
            if error.code() == windows::Win32::Foundation::ERROR_FILE_NOT_FOUND.to_hresult() =>
        {
            Ok(())
        }
        Err(_) => Err(Refused::NoWrite),
    }
}

/// Start the task now, which is how the app comes back elevated after the row was turned on: the
/// very record that starts it at a logon is the one that starts it here, so what a user gets is
/// what they will get tomorrow rather than a second way of starting it.
pub(crate) fn run_now() -> Result<(), Refused> {
    let service = connect().ok_or(Refused::NoWrite)?;

    // Safety: as above; `Run` takes the parameters a task may be given, which this one has none of.
    unsafe {
        let root = service
            .GetFolder(&BSTR::from("\\"))
            .map_err(|_| Refused::NoWrite)?;
        let task = root
            .GetTask(&BSTR::from(TASK_NAME))
            .map_err(|_| Refused::NoWrite)?;

        // The parameters a task may be given, of which this one is handed none — a `Run` value has
        // no arguments either, which is the whole of what the task replaces.
        task.Run(&VARIANT::default())
            .map_err(|_| Refused::NoWrite)?;
    }

    Ok(())
}

/// The Task Scheduler's own object, connected to this machine.
///
/// **The apartment is the process's, and it is made here rather than trusted to the caller.** A
/// thread that has never joined one cannot create anything: `CoCreateInstance` answers
/// `CO_E_NOTINITIALIZED` (measured: `0x800401F0`), and the two callers that write a task are exactly
/// such threads — the elevated helper makes its one write and exits before the tray's own
/// `CoInitializeEx`, and the reconciler at a start runs ahead of it too. A write that could not even
/// be attempted was reported as the machine refusing one, which is the one answer a user cannot act
/// on. `CoIncrementMTAUsage` is the platform's own answer for code that cannot control its caller:
/// it puts the process in an apartment without moving the calling thread, and a thread that has used
/// it can still join one of its own afterwards (measured: `S_OK`). The count is never given back —
/// the process is what holds it, and this app's process lives until it is closed.
fn connect() -> Option<ITaskService> {
    static APARTMENT: OnceLock<()> = OnceLock::new();

    APARTMENT.get_or_init(|| {
        // Safety: a call of the platform's with no arguments; the cookie it answers is the
        // process's and is deliberately not held, because it is never given back.
        let _ = unsafe { CoIncrementMTAUsage() };
    });

    // Safety: a local server object of the Shell's own; the apartment is the process's, made above.
    unsafe {
        let service: ITaskService =
            CoCreateInstance(&TaskScheduler, None, CLSCTX_INPROC_SERVER).ok()?;
        service
            .Connect(
                &VARIANT::default(),
                &VARIANT::default(),
                &VARIANT::default(),
                &VARIANT::default(),
            )
            .ok()?;
        Some(service)
    }
}

/// The task document: the whole of what [`register`] writes.
///
/// `user` is written into both the trigger and the principal, which is what makes the task belong
/// to one account rather than to whoever happens to sign in. Both are escaped, because a path and
/// an account name are text a person's machine provided rather than text this app decided.
pub(crate) fn document(exe: &Path, user: &str) -> String {
    let command = escape(&crate::paths::plain_path(exe));
    let working = escape(
        &exe.parent()
            .map(crate::paths::plain_path)
            .unwrap_or_else(|| "C:\\".to_string()),
    );

    format!(
        r#"<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>Rust Hover Preview</Author>
    <Description>{description}</Description>
    <URI>\{name}</URI>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{user}</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{user}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{command}</Command>
      <WorkingDirectory>{working}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"#,
        description = escape(TASK_DESCRIPTION),
        name = TASK_NAME,
        // The account is escaped where it is written rather than into a binding, because a format
        // string's own braces are the other thing this document holds: the description is escaped
        // into its argument the same way round.
        user = escape(user),
    )
}

/// Text as an XML element's own content: the five characters a document cannot hold as they are.
fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());

    for character in text.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            other => out.push(other),
        }
    }

    out
}

/// Whether two spellings are the path of one file, as the machine reads a path.
///
/// The same question `app::startup` asks of a `Run` value, and answered the same way: the spelling
/// settles what it can, and where the two read differently the files themselves are asked, which is
/// what settles a short name for a folder or a junction.
fn same_path(registered: &str, running: &Path) -> bool {
    let left = crate::paths::plain_path(Path::new(registered.trim()));
    let right = crate::paths::plain_path(running);

    if left.eq_ignore_ascii_case(&right) {
        return true;
    }

    match (std::fs::canonicalize(&left), std::fs::canonicalize(&right)) {
        (Ok(left), Ok(right)) => {
            crate::paths::plain_path(&left).eq_ignore_ascii_case(&crate::paths::plain_path(&right))
        }
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The document this app writes, read back: every fact the tray draws a row from survives the
    /// round trip, and the two things a logon needs — the run level and the account — are in it.
    #[test]
    fn the_task_this_app_writes_is_read_back_as_it_was_written() {
        let exe = Path::new(r"C:\Program Files\Rust Hover Preview\rust-hover-preview.exe");
        let xml = document(exe, r"DESKTOP\someone");

        let facts = facts_of_xml(&xml).expect("the document this app writes is one it can read");

        assert!(facts.present, "a document is a task");
        assert!(facts.enabled, "a logon may run it");
        assert!(
            facts.elevated,
            "highest available is the whole point: a logon starts it as an administrator"
        );
        assert_eq!(
            facts.command.as_deref(),
            Some(r"C:\Program Files\Rust Hover Preview\rust-hover-preview.exe"),
            "the action is the exe's path, unquoted, because a path is not a command line"
        );
        assert!(
            facts.starts_this_copy(exe),
            "the task names the copy it was written for"
        );
        assert!(
            !facts.starts_this_copy(Path::new(
                r"C:\Program Files\Other Folder\rust-hover-preview.exe"
            )),
            "and not another copy of the app"
        );
    }

    /// The account is written into both halves of the document, so a task belongs to the logon it
    /// was made for rather than to whoever signs in next.
    #[test]
    fn the_account_is_in_the_trigger_and_in_the_principal() {
        let xml = document(Path::new(r"C:\app\app.exe"), r"LAPTOP\name");

        assert_eq!(
            xml.matches(r"<UserId>LAPTOP\name</UserId>").count(),
            2,
            "the trigger and the principal both name the account:\n{xml}"
        );
    }

    /// A path holding the characters a document cannot carry is carried anyway, and read back as
    /// the path it is: a user's folder is theirs to name.
    #[test]
    fn a_path_of_odd_characters_survives_the_document() {
        let exe = Path::new(r"C:\a & b<c>d\app.exe");
        let xml = document(exe, "machine & user");

        assert!(
            !xml.contains("a & b<c>d"),
            "the raw characters are escaped in the document"
        );

        let facts = facts_of_xml(&xml).expect("a document of escaped text is a document");
        assert!(
            facts.starts_this_copy(exe),
            "and reads back as the same path"
        );
    }

    /// A task somebody turned off in Task Scheduler starts nothing, and the row has to be able to
    /// say so: the flag is read from the document exactly like a `Run` value's approved flag.
    #[test]
    fn a_disabled_task_is_not_one_a_logon_runs() {
        let xml = document(Path::new(r"C:\app\app.exe"), "user").replace(
            "<Enabled>true</Enabled>\n      <UserId>",
            "<Enabled>false</Enabled>\n      <UserId>",
        );

        let facts = facts_of_xml(&xml).expect("still a document");
        assert!(
            !facts.enabled,
            "a disabled trigger is a task that runs nothing"
        );
        assert!(
            !facts.starts_this_copy(Path::new(r"C:\app\app.exe")),
            "and it is not a machine that starts this app"
        );

        let xml = document(Path::new(r"C:\app\app.exe"), "user")
            .replace("<Enabled>true</Enabled>", "<Enabled>false</Enabled>");
        let facts = facts_of_xml(&xml).expect("still a document");
        assert!(!facts.enabled, "so is a disabled task");
    }

    /// A run level of `LeastPrivilege` is the task this app must not leave behind: it starts the
    /// app, unelevated, which is exactly the state the row exists to leave.
    #[test]
    fn a_task_that_runs_at_the_lowest_level_is_not_an_elevated_start() {
        let xml = document(Path::new(r"C:\app\app.exe"), "user")
            .replace("HighestAvailable", "LeastPrivilege");

        let facts = facts_of_xml(&xml).expect("still a document");
        assert!(!facts.elevated);
        assert!(!facts.starts_this_copy(Path::new(r"C:\app\app.exe")));
    }

    /// Text that is not a task document is not a task: the reader answers nothing rather than
    /// making one out of a file it does not understand — and an XML document that is not a `Task` is
    /// not one either, however many of the elements this reads it happens to hold.
    #[test]
    fn text_that_is_not_a_task_document_is_not_read_as_one() {
        assert!(facts_of_xml("").is_none());
        assert!(facts_of_xml("not xml at all").is_none());
        assert!(facts_of_xml("<Task>").is_none());
        assert!(
            facts_of_xml(
                "<NotATask><RunLevel>HighestAvailable</RunLevel><Command>C:\\app.exe</Command></NotATask>"
            )
            .is_none(),
            "a document of the right elements under the wrong root is not a task"
        );
    }

    /// The Task Scheduler is reachable from a thread that never asked for an apartment of its own,
    /// which is the thread the elevated helper makes its one write on. `CoCreateInstance` answers
    /// `CO_E_NOTINITIALIZED` there, and a write that could not even be attempted used to be reported
    /// as the machine refusing one — a dialog naming a security policy over a thread that had never
    /// asked COM for anything (see `connect`).
    #[test]
    fn the_task_service_is_reachable_from_a_thread_with_no_apartment_of_its_own() {
        let reached = std::thread::spawn(|| connect().is_some())
            .join()
            .expect("the probe thread answered");

        assert!(
            reached,
            "a thread that has never called CoInitializeEx can still reach the Task Scheduler"
        );
    }

    /// The document is one the Task Scheduler itself takes: `SetXmlText` parses it, so every element
    /// sits where the service expects it and no value is out of range. That parser is the one a task
    /// has to satisfy — the reader below can only check what it was written to check, and a document
    /// only this app can read is a task no logon ever runs.
    #[test]
    fn the_document_is_one_the_task_scheduler_itself_takes() {
        let exe = Path::new(r"C:\Program Files\Rust Hover Preview\rust-hover-preview.exe");
        let service = connect().expect("the Task Scheduler answers");

        // Safety: a definition of the service's own, and a document this app wrote.
        let accepted = unsafe {
            let definition = service.NewTask(0).expect("a new definition");
            definition.SetXmlText(&BSTR::from(document(exe, r"DESKTOP\someone")))
        };

        assert!(
            accepted.is_ok(),
            "the service parses the document: {accepted:?}"
        );
    }

    /// The service's own copy of a task is UTF-16, and is read with or without its byte order mark.
    #[test]
    fn the_services_own_copy_of_a_task_is_utf16() {
        let text = document(Path::new(r"C:\app\app.exe"), "user");
        let units: Vec<u8> = text.encode_utf16().flat_map(u16::to_le_bytes).collect();

        let marked: Vec<u8> = [0xFF, 0xFE]
            .into_iter()
            .chain(units.iter().copied())
            .collect();

        assert_eq!(decode_utf16(&marked).as_deref(), Some(text.as_str()));
        assert_eq!(decode_utf16(&units).as_deref(), Some(text.as_str()));
    }
}
