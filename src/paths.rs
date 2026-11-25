//! Paths in the spelling the Shell will take — and the one folder a run keeps its own files in.
//!
//! The Explorer hook canonicalizes to the verbatim `\\?\` form, which is not a legal
//! thing to hand a Shell call, a browser URL or a media source: a file named in that
//! form is a file none of them can find. The prefix comes off before anything is
//! pointed at a file, and a share keeps its server.

use std::path::{Path, PathBuf};
use std::sync::OnceLock;

/// The folder this run keeps `config.ini`, `theme`, `engines`, `webview`, `libreoffice` and
/// `.env.local` in: the exe's own folder where a `.portable` marker stands beside it — kept in
/// `rust-hover-preview` there, so an unzip does not scatter the app's files among the user's —
/// and the machine's own folders otherwise.
///
/// One root for all of them rather than a folder per kind, because what a portable build promises
/// is that nothing of it is left on the machine — and a folder split across `%APPDATA%` and
/// `%LOCALAPPDATA%` is a promise kept by six call sites agreeing instead of by one.
///
/// A test build never reads the machine's folders here: each of the six keeps its own test folder,
/// and every one of them answers from that before asking this at all.
pub(crate) fn data_root() -> Option<&'static Path> {
    static ROOT: OnceLock<Option<PathBuf>> = OnceLock::new();
    ROOT.get_or_init(portable_root).as_deref()
}

#[cfg(test)]
fn portable_root() -> Option<PathBuf> {
    None
}

#[cfg(not(test))]
fn portable_root() -> Option<PathBuf> {
    let dir = std::env::current_exe().ok()?.parent()?.to_path_buf();
    // `.portable` beside the exe, the way a portable editor says so: its presence is the whole of
    // what it says, whatever it holds.
    dir.join(".portable")
        .is_file()
        .then(|| dir.join("rust-hover-preview"))
}

/// Whether this run is the portable one: a `.portable` file beside the exe, the way a portable
/// editor says so.
pub(crate) fn portable() -> bool {
    data_root().is_some()
}

/// A path with the Shell's verbatim prefix taken off.
pub(crate) fn plain_path(path: &Path) -> String {
    let text = path.to_string_lossy();

    match text.strip_prefix(r"\\?\UNC\") {
        Some(rest) => format!(r"\\{rest}"),
        None => text.strip_prefix(r"\\?\").unwrap_or(&text).to_string(),
    }
}

/// The same path with its separators turned round: `E:/Videos/a b.mkv`.
///
/// Nothing is escaped — a program that reads `/` but treats `\` as an escape (`ffmpeg`'s own
/// option parser, most `-`-flagged tools written for both platforms) is the one this is for, and
/// it takes the name as a name rather than as a URL.
pub(crate) fn forward_slashes(path: &Path) -> String {
    plain_path(path).replace('\\', "/")
}

/// The path in the form the volumes' own 8.3 names give it — `E:\VIDEO~1\A~1.MKV` — where it has
/// one.
///
/// `None` for a path that does not exist, for a volume where 8.3 names are turned off (Windows
/// 11's default on a non-system volume), and for a name that is already an 8.3 name: the point of
/// asking is to be handed a *different* spelling, and the caller's own fallback is the plain path.
pub(crate) fn short_path(path: &Path) -> Option<String> {
    use windows::core::PCWSTR;
    use windows::Win32::Storage::FileSystem::GetShortPathNameW;

    let plain = plain_path(path);
    let wide: Vec<u16> = plain.encode_utf16().chain(std::iter::once(0)).collect();
    let mut asked = 0u32;
    let mut buffer = vec![0u16; 260];
    loop {
        let answered = unsafe { GetShortPathNameW(PCWSTR(wide.as_ptr()), Some(&mut buffer)) };
        // The answer is the length written, or — where the buffer was short — the length needed,
        // NUL excluded either way.
        if answered == 0 {
            return None;
        }
        let held = answered as usize;
        if held < buffer.len() {
            let short = String::from_utf16_lossy(&buffer[..held]);
            return (short != plain).then_some(short);
        }
        if asked >= 2 {
            return None;
        }
        asked += 1;
        buffer = vec![0u16; held + 1];
    }
}

/// A local absolute path as a `file://` URL the way a program that takes a URL reads it:
/// `file:///E:/Videos/a%20b.mkv`, and `file://server/share/a.mkv` for a share.
///
/// `None` for a path that is not absolute, which is not a URL anything can open.
pub(crate) fn file_url(path: &Path) -> Option<String> {
    file_url_of(path, true)
}

/// The same URL with nothing escaped: `file:///E:/Videos/a b.mkv`.
///
/// For the odd program that takes a URL and then hands `%20` through to the filesystem as those
/// four characters. The separators are still slashes, because that is what makes it a URL.
pub(crate) fn file_url_literal(path: &Path) -> Option<String> {
    file_url_of(path, false)
}

fn file_url_of(path: &Path, escape: bool) -> Option<String> {
    if !path.is_absolute() {
        return None;
    }

    let local = plain_path(path);

    let (mut url, rest) = match local.strip_prefix(r"\\") {
        Some(share) => (String::from("file://"), share.to_string()),
        None => (String::from("file:///"), local),
    };

    for character in rest.chars() {
        match character {
            '\\' => url.push('/'),
            ' ' if escape => url.push_str("%20"),
            '#' if escape => url.push_str("%23"),
            '?' if escape => url.push_str("%3F"),
            '%' if escape => url.push_str("%25"),
            other => url.push(other),
        }
    }

    Some(url)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_run_with_no_portable_marker_keeps_the_machines_own_folders() {
        // Every test answers from a folder of its own before this is asked at all, and the marker
        // is a file beside the exe a test build does not put one beside.
        assert!(data_root().is_none());
        assert!(!portable());
    }

    #[test]
    fn a_verbatim_path_is_turned_round_shortened_and_written_as_a_url() {
        // Separators, with the verbatim prefix off so the two spellings are the same file.
        assert_eq!(
            forward_slashes(Path::new(r"\\?\E:\Videos\a b.mkv")),
            "E:/Videos/a b.mkv"
        );
        assert_eq!(
            forward_slashes(Path::new(r"\\server\share\a.mkv")),
            "//server/share/a.mkv"
        );

        // A URL escapes what a URL may not carry, and a share keeps its server.
        assert_eq!(
            file_url(Path::new(r"\\?\E:\Videos\a b.mkv")).as_deref(),
            Some("file:///E:/Videos/a%20b.mkv")
        );
        assert_eq!(
            file_url(Path::new(r"\\server\share\a b.mkv")).as_deref(),
            Some("file://server/share/a%20b.mkv")
        );
        // The literal form is the same URL with the escaping left out.
        assert_eq!(
            file_url_literal(Path::new(r"E:\Videos\a b.mkv")).as_deref(),
            Some("file:///E:/Videos/a b.mkv")
        );
        // A relative path is not a URL anything can open.
        assert_eq!(file_url(Path::new(r"Videos\a.mkv")), None);
        // A path that is not there has no 8.3 name to be handed.
        assert_eq!(
            short_path(Path::new(r"C:\no such folder\no such file.mkv")),
            None
        );
    }

    /// The 8.3 form where the volume keeps one, and nothing where it does not: a machine with
    /// 8.3 names turned off answers the plain path back, and that is not a spelling to hand a
    /// program as if it were a short one.
    #[test]
    fn a_short_name_is_a_different_spelling_of_the_same_file() {
        let folder = std::env::temp_dir().join("rhp short name probe");
        let _ = std::fs::create_dir_all(&folder);
        let file = folder.join("a long enough file name.mkv");
        if std::fs::write(&file, b"x").is_err() {
            return;
        }
        match short_path(&file) {
            Some(short) => {
                assert!(Path::new(&short).is_file(), "{short} is the same file");
                assert!(
                    short.len() < plain_path(&file).len(),
                    "{short} is a shorter spelling than the name it stands for"
                );
            }
            None => assert_eq!(
                short_path(&file).map(|short| short == plain_path(&file)),
                None,
                "a volume without 8.3 names answers nothing rather than the same name back"
            ),
        }
        let _ = std::fs::remove_dir_all(&folder);
    }

    #[test]
    fn a_verbatim_path_is_read_as_the_spelling_the_shell_will_take() {
        assert_eq!(
            plain_path(Path::new(r"\\?\C:\docs\report.docx")),
            r"C:\docs\report.docx"
        );
        assert_eq!(
            plain_path(Path::new(r"\\?\UNC\server\share\report.docx")),
            r"\\server\share\report.docx"
        );
        // A path that was never verbatim is already the right spelling.
        assert_eq!(
            plain_path(Path::new(r"C:\docs\report.docx")),
            r"C:\docs\report.docx"
        );
        assert_eq!(
            plain_path(Path::new(r"\\server\share\track.mp3")),
            r"\\server\share\track.mp3"
        );
    }
}
