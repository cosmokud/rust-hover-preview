//! Which account this process runs as, and whether its token is the elevated half of it.
//!
//! Two things in this app are asked of the machine rather than of the app, and both of them
//! are answered differently by an elevated process: the startup entry, which belongs to the
//! account whose **logon** starts this app, and the browser engine, whose own process refuses
//! to run on an elevated token. Neither can be reasoned about from the process alone, so what
//! they need is here: the SID of the account this process runs as, the SID of the account the
//! **desktop** belongs to, and whether the process's token is elevated.
//!
//! **The two SIDs are not always one account, and that is the whole point of the comparison.**
//! An administrator who answer UAC with their own credentials runs the elevated half of their
//! own account: one SID, one `HKEY_CURRENT_USER`, one `Run` key, and the entry an elevated run
//! writes is the entry the next logon reads. A standard user whose elevation prompt is answered
//! by an *administrator's* credentials — the over-the-shoulder prompt Windows shows on a managed
//! machine, and what "run as administrator" does when it asks for a different name — runs as that
//! administrator: `HKEY_CURRENT_USER` is then the administrator's hive, an entry written there
//! belongs to an account that will never be logged on to, and the desktop's own logon starts
//! nothing at all. The desktop's account is read from the Shell's own token rather than guessed
//! at, because `explorer.exe` is the one process in the session that is certainly the desk the
//! user is sitting at.
//!
//! Reading the Shell's token is the one crossing here that can fail, and it is allowed in both
//! directions: a higher integrity process may open a lower one's token, and a medium-integrity
//! process may not read a higher one's — which is why nothing here is asked of a run that is not
//! elevated. An answer that cannot be had is `None`, and every caller treats that as "the
//! ordinary case", which is the one where nothing has to be redirected.

use std::mem::size_of;
use std::sync::OnceLock;

use windows::core::PWSTR;
use windows::Win32::Foundation::{CloseHandle, LocalFree, HANDLE, HLOCAL};
use windows::Win32::Security::Authorization::ConvertSidToStringSidW;
use windows::Win32::Security::{
    GetTokenInformation, TokenElevation, TokenUser, PSID, TOKEN_ELEVATION, TOKEN_QUERY, TOKEN_USER,
};
use windows::Win32::System::Threading::{
    GetCurrentProcess, OpenProcess, OpenProcessToken, PROCESS_QUERY_LIMITED_INFORMATION,
};
use windows::Win32::UI::WindowsAndMessaging::{GetShellWindow, GetWindowThreadProcessId};

/// Whether this process's token is an elevated one.
///
/// Asked once and kept: a token's elevation cannot change while the process lives, and this is
/// read on the start's own path — before the tray icon — and again by every browser this app
/// begins.
pub(crate) fn is_elevated() -> bool {
    static ELEVATED: OnceLock<bool> = OnceLock::new();

    *ELEVATED.get_or_init(|| token_is_elevated(unsafe { GetCurrentProcess() }).unwrap_or(false))
}

/// The SID of the account the desktop belongs to — the Shell's own — or nothing where there is
/// no Shell to ask, or where a token this process may not read is in the way.
pub(crate) fn desktop_user_sid() -> Option<String> {
    static SID: OnceLock<Option<String>> = OnceLock::new();

    SID.get_or_init(|| {
        // Safety: `GetShellWindow` answers a window or null, and the pid is read out of the
        // window rather than out of anything this process holds.
        let pid = unsafe {
            let shell = GetShellWindow();
            if shell.is_invalid() {
                return None;
            }
            let mut pid = 0u32;
            GetWindowThreadProcessId(shell, Some(&mut pid));
            pid
        };

        if pid == 0 {
            return None;
        }

        // The Shell is a process of another program's, so it is asked for the least a token can
        // be opened through — and it is opened, not inherited: a handle this function made is a
        // handle this function closes.
        let process = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) }.ok()?;
        let sid = user_sid_of(process);
        unsafe {
            let _ = CloseHandle(process);
        }
        sid
    })
    .clone()
}

/// The SID of the account this process runs as.
pub(crate) fn our_user_sid() -> Option<String> {
    static SID: OnceLock<Option<String>> = OnceLock::new();

    SID.get_or_init(|| user_sid_of(unsafe { GetCurrentProcess() }))
        .clone()
}

/// Whether this process runs as an account other than the one the desktop belongs to.
///
/// An answer that could not be had on either side is `false`: what this decides is whether the
/// startup entry has to be written into another hive, and the ordinary run — where the entry is
/// the running account's own — is the one to be wrong about.
pub(crate) fn runs_under_another_account() -> bool {
    !same_account(&desktop_user_sid(), &our_user_sid())
}

/// Whether two answers name one account, where an answer that is missing is not a difference.
fn same_account(left: &Option<String>, right: &Option<String>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left.eq_ignore_ascii_case(right),
        _ => true,
    }
}

/// The user account a process's token names, as the string the Shell spells it with.
fn user_sid_of(process: HANDLE) -> Option<String> {
    let token = open_token(process)?;
    let sid = token_user_sid(token);
    // Safety: the handle `open_token` made and nothing else holds.
    unsafe {
        let _ = CloseHandle(token);
    }
    sid
}

fn open_token(process: HANDLE) -> Option<HANDLE> {
    let mut token = HANDLE::default();
    // Safety: a handle is written into a local that outlives the call, and it is closed by the
    // caller.
    unsafe { OpenProcessToken(process, TOKEN_QUERY, &mut token) }.ok()?;
    Some(token)
}

/// One token's `TokenUser` SID, as text.
///
/// The information classes are asked twice, the way every `GetTokenInformation` is: once with no
/// buffer for the size it needs — which fails, and is what asks — and once with a buffer that
/// size.
fn token_user_sid(token: HANDLE) -> Option<String> {
    let mut needed = 0u32;
    // Safety: no buffer is offered, and the failure that answers it is the one expected.
    let _ = unsafe { GetTokenInformation(token, TokenUser, None, 0, &mut needed as *mut u32) };

    if needed == 0 {
        return None;
    }

    let mut buffer = vec![0u8; needed as usize];
    // Safety: the buffer is `needed` bytes, which is what the call above asked for.
    unsafe {
        GetTokenInformation(
            token,
            TokenUser,
            Some(buffer.as_mut_ptr().cast()),
            needed,
            &mut needed as *mut u32,
        )
    }
    .ok()?;

    // Safety: a `TOKEN_USER` whose `SID_AND_ATTRIBUTES::Sid` points into the buffer the call
    // just filled, which outlives this read.
    let user: TOKEN_USER = unsafe { std::ptr::read_unaligned(buffer.as_ptr().cast()) };

    sid_as_text(user.User.Sid)
}

/// A SID as the string `S-1-5-21-…` the Shell spells it with.
fn sid_as_text(sid: PSID) -> Option<String> {
    if sid.is_invalid() {
        return None;
    }

    let mut text = PWSTR::null();
    // Safety: a `PWSTR` is written into a local, and the string it names is allocated by the API
    // and freed below.
    unsafe { ConvertSidToStringSidW(sid, &mut text) }.ok()?;

    // Safety: the string `ConvertSidToStringSidW` made, read once.
    let spelled = unsafe { text.to_string().ok() };

    // Safety: a `LocalAlloc`'d string is the caller's to free, whether or not it was read.
    unsafe {
        let _ = LocalFree(HLOCAL(text.0.cast()));
    }

    spelled
}

fn token_is_elevated(process: HANDLE) -> Option<bool> {
    let token = open_token(process)?;

    let mut elevation = TOKEN_ELEVATION::default();
    let mut needed = 0u32;
    // Safety: a `TOKEN_ELEVATION` is offered, which is the size said beside it, and the token is
    // closed after the answer has been read out of it.
    let asked = unsafe {
        GetTokenInformation(
            token,
            TokenElevation,
            Some((&mut elevation as *mut TOKEN_ELEVATION).cast()),
            size_of::<TOKEN_ELEVATION>() as u32,
            &mut needed as *mut u32,
        )
    };
    // Safety: the handle `open_token` made and nothing else holds.
    unsafe {
        let _ = CloseHandle(token);
    }

    asked.ok()?;
    Some(elevation.TokenIsElevated != 0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_answer_that_cannot_be_had_is_not_another_account() {
        // A missing SID on either side reads as the ordinary run rather than as a redirect: the
        // startup entry the running account can write is the safe wrong answer, and another
        // account's hive is not. Two answers that are the same string are one account, in
        // whatever case the Shell spelled them.
        assert!(same_account(
            &Some("S-1-5-21-1".into()),
            &Some("S-1-5-21-1".into())
        ));
        assert!(same_account(
            &Some("S-1-5-21-1".into()),
            &Some("s-1-5-21-1".into())
        ));
        // And an answer that is missing on either side is not a difference to act on.
        assert!(same_account(&None, &Some("S-1-5-21-1".into())));
        assert!(same_account(&Some("S-1-5-21-1".into()), &None));
        assert!(!same_account(
            &Some("S-1-5-21-1".into()),
            &Some("S-1-5-21-2".into())
        ));
    }

    #[test]
    fn elevation_is_answered_once_and_kept() {
        // The token cannot change while the process lives, so the two readings are one answer —
        // and reading it here must not be the thing that fails a test on a machine whose token
        // this process may not open.
        assert_eq!(is_elevated(), is_elevated());
    }
}
