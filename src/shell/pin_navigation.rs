//! The list a pinned window's own previous/next buttons step through, and the order it is in.
//!
//! A pin that is up is a window of its own, and two of its caption buttons are a way through
//! the folder it was taken up in: the file before this one and the file after it, in the order
//! the listing is showing them. What this module is that list, and the one thing it is
//! expensive about is the folder a pin sits in — a folder of a thousand files is a folder this
//! app must not open a thousand of.
//!
//! **What a scan is allowed to do.** The directory read is one flat `read_dir`, with no
//! recursion and no file opened. The filter is the name and nothing else: which kind claims a
//! name (`routing::kind_of`) and whether a reader is installed for it
//! (`routing::readers_for`). It is deliberately *not* `is_media_file_with_facts`, which
//! sniffs a file's header by opening it — on a large folder that is a thousand opens and a
//! thousand downloads if the folder is on OneDrive.
//!
//! **What a scan is not allowed to ask of a file.** `DirEntry::file_type()` is data the OS
//! returned with the name anyway (see `document_cache`, which reads a folder the same way),
//! and a reparse point is not asked again: `metadata()` on one of those hydrates a OneDrive
//! placeholder and downloads the file — or, when a folder is full of them, the folder. Where
//! the sort is by name, nothing is asked of a file at all.
//!
//! **What the scan is remembered as.** The map holds each entry's *kind* beside its path, so
//! switching `All` to `Category` re-filters what is already in hand rather than reading the
//! folder again. It is capped, because a hand that pins files in a thousand folders should
//! not leave a thousand folders behind it; a pin in a folder not in the map reads it again,
//! which is the cost of the cap and the whole of it.

use crate::config::config::{AppConfig, PinNavFileTypes, PreviewType};
use crate::formats::routing;
use crate::shell::explorer_hook::{view_sort_of, SortKey, ViewSort};
use once_cell::sync::Lazy;
use std::cmp::Ordering;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;

/// Every folder a pin has been asked to walk, and the files in each that a pin could be
/// shown. One scan each, and a scan is a name list rather than a read of the files.
static FOLDERS: Lazy<Mutex<HashMap<PathBuf, FolderEntries>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// How many folders are kept. A pin walks one folder at a time, so this is generous for
/// what the app does and small enough that a machine swept across a whole drive does not
/// grow the map by a folder per window.
const FOLDER_LIMIT: usize = 32;

/// What one folder holds for the walk: the files, already in the order the listing was
/// showing them.
struct FolderEntries {
    entries: Vec<NavEntry>,
}

/// One file of the walk: where it is, what kind claimed it, and — only where the order needs
/// them — the two facts of the file itself the order is read from. A name-ordered walk never
/// fills them in, and that is the point: the walk is a list of names, and only a folder
/// sorted by date or size is worth asking a file anything.
#[derive(Clone)]
struct NavEntry {
    path: PathBuf,
    kind: PreviewType,
    facts: Option<Facts>,
}

/// The files a pin in `current`'s folder steps through, in the order it should step, or
/// nothing where the file is one no kind claims or the folder cannot be read.
///
/// The list is the folder filtered and sorted, then filtered again by the switch that says
/// whether "next" means the next file or the next file of the same kind of thing. Both
/// answers are read from one scan: the kinds are held per entry, so `Category` is a
/// re-filter of what is already in hand rather than a second walk of the folder.
///
/// The order is the listing's own where the listing has one to give (see [`ViewSort`]) and
/// natural name order where it does not — which is also what an Explorer default view
/// shows, since `Name` ascending *is* the natural comparison a file manager uses.
pub fn list_for(current: &Path, config: &AppConfig) -> Option<Vec<PathBuf>> {
    let folder = current.parent()?;

    let entries = entries_for(folder, config)?;

    let mode = config.pin_nav_file_types;
    let wanted = routing::kind_of(current, config).map(routing::nav_category);

    Some(
        entries
            .iter()
            .filter(|entry| {
                mode == PinNavFileTypes::All
                    || (wanted.is_some()
                        && wanted == Some(routing::nav_category(entry.kind)))
            })
            .map(|entry| entry.path.clone())
            .collect(),
    )
}

/// The file a step of `step` from `current` lands on, given the list the walk is made of.
///
/// * A file in the list steps, and the ends wrap: the first file's previous is the last, the
///   last file's next is the first. A folder is a ring, and a pin at either end of it is at
///   the end rather than stuck there.
/// * A file *not* in the list is the case a switch can put it in: `Category` narrows the
///   walk, and a file of another category is not on it. There is no position to step from,
///   so the step is taken from the end it points at — `Next` goes to the first file, which is
///   the nearest thing to where the user is looking, and `Previous` to the last, which is
///   the same walk run the other way. The pinned file is not lost: it is the file the pin
///   was on, and it is the file the caption still names until the next click.
pub fn step_to(current: &Path, list: &[PathBuf], step: i32) -> Option<PathBuf> {
    if list.is_empty() {
        return None;
    }

    let Some(position) = list.iter().position(|entry| same_file(entry, current)) else {
        return if step >= 0 {
            list.first().cloned()
        } else {
            list.last().cloned()
        };
    };

    let last = list.len() as i32 - 1;
    let from = position as i32;
    // A step of zero is a step onto the file already there, which is what a caller asking
    // for no movement should get rather than a panic on the modulo.
    let next = (from + step).rem_euclid(last + 1);
    list.get(next as usize).cloned()
}

/// The folder's own entries, read on a miss and remembered after. The sort is read from what
/// the view drawing the folder last said, and it is read *before* the files so that a walk
/// is ordered by what the listing was showing when the pin went up rather than by what it
/// shows now.
fn entries_for(folder: &Path, config: &AppConfig) -> Option<Vec<NavEntry>> {
    if let Ok(folders) = FOLDERS.lock() {
        if let Some(cached) = folders.get(folder) {
            return Some(cached.entries.clone());
        }
    }

    let sort = view_sort_of(folder);
    let entries = scan(folder, config, sort);
    let ordered = {
        let mut entries = entries;
        order(&mut entries, sort);
        entries
    };

    let Ok(mut folders) = FOLDERS.lock() else {
        return Some(ordered);
    };

    if folders.len() >= FOLDER_LIMIT {
        // Whichever folder the map offers up, rather than the one least recently walked. A
        // walk is a name list and costs the same whichever folder it is a list of, so what
        // is given up here is a list that will be read again, not anything already done.
        if let Some(dropped) = folders.keys().next().cloned() {
            folders.remove(&dropped);
        }
    }
    folders.insert(
        folder.to_path_buf(),
        FolderEntries {
            entries: ordered.clone(),
        },
    );

    Some(ordered)
}

/// What a walk of one folder costs, and what it is not allowed to cost.
///
/// The engine check is the expensive part and it is the part that has to be hoisted: whether
/// LibreOffice, ImageMagick, PeaZip or Calibre is on this machine is one question, and asking
/// it per file would be a registry lookup per `.docx` in the folder. It is asked once per
/// kind that turns up, and the answer is kept for the rest of the scan.
fn scan(folder: &Path, config: &AppConfig, sort: Option<ViewSort>) -> Vec<NavEntry> {
    let Ok(read) = std::fs::read_dir(folder) else {
        return Vec::new();
    };

    // Whether a sort needs a file's own metadata at all: name and type are read off the
    // name, and only date and size are worth asking a file for.
    let needs_metadata = matches!(
        sort.map(|sort| sort.key),
        Some(SortKey::DateModified | SortKey::Size)
    );

    let mut installed: HashMap<PreviewType, bool> = HashMap::new();
    let mut entries: Vec<NavEntry> = Vec::new();

    for entry in read.flatten() {
        // A file's own kind, from the directory entry the OS already returned. `file_type`
        // does not open the file and does not follow a reparse point, so a shortcut, a
        // folder, a OneDrive placeholder and a symbolic link are all skipped without a
        // single byte of their content being read (see `document_cache`, which walks a
        // folder this same way).
        if !entry.file_type().is_ok_and(|kind| kind.is_file()) {
            continue;
        }

        let path = entry.path();

        // The name gates the kind and the kind gates everything else. A name no list claims
        // is not a file this app previews, and the walk stops there.
        //
        // Asked of the name alone rather than of the file: this is the one place in the app
        // that walks a folder rather than one file, and the answer this claim gives for a
        // handful of names — a `.ts` a file could be told from a name — is a read of every
        // file that carries one. On a folder of thousands that is thousands of opens, and on
        // a cloud folder it is the whole folder downloaded.
        //
        // What it costs is that a transport stream is walked as the TypeScript it shares its
        // name with. That is the one answer this walk gives that a hover would not, and it is
        // the right way round: a file walked onto is loaded whole, where the content answers
        // for itself.
        let Some(kind) = routing::kind_of_name(&path, config) else {
            continue;
        };

        // A kind the tray has switched off is not walked, whatever else it is. This is
        // asked per entry rather than cached: the gate is a field read, and it can be
        // thrown between two steps of the same pin.
        if !kind.enabled_in(config) {
            continue;
        }

        // Whether a reader is on the machine is one question per kind, asked once and
        // remembered — `readers_for` narrows a kind's chain to the readers this machine has,
        // and it asks the registry behind that answer (see `office_formats::app_installed`).
        let can_read = match installed.get(&kind) {
            Some(known) => *known,
            None => {
                let can_read = !routing::readers_for(kind, &path).is_empty();
                installed.insert(kind, can_read);
                can_read
            }
        };
        if !can_read {
            continue;
        }

        // Metadata is read only where the order needs it, and only off a file the OS has
        // already said is a file — never off a reparse point, which is the difference
        // between reading a date and downloading a placeholder.
        if needs_metadata {
            if let Ok(metadata) = entry.metadata() {
                entries.push(NavEntry {
                    path,
                    kind,
                    facts: Some(Facts {
                        modified: metadata.modified().ok(),
                        bytes: metadata.len(),
                    }),
                });
                continue;
            }
        }

        entries.push(NavEntry {
            path,
            kind,
            facts: None,
        });
    }

    entries
}

/// The two facts a date or size order needs, kept beside the entry so the sort does not have
/// to ask a file again. `None` for a name-ordered walk, which asks nothing.
#[derive(Clone)]
struct Facts {
    modified: Option<SystemTime>,
    bytes: u64,
}

/// The order a listing's files are stepped through in.
///
/// The default is natural name order, which is not the same as byte order and is what
/// Explorer shows for `Name` ascending: a folder of exported frames holds `img2.png` before
/// `img10.png`, and a byte order puts `img10` first because `1` sorts before `2`. So a
/// natural comparison is what a walk with no sort of its own gets, and it is also what a
/// `Name` sort gets — the same answer, reached by the same rule.
fn order(entries: &mut [NavEntry], sort: Option<ViewSort>) {
    let key = sort.map(|sort| sort.key).unwrap_or(SortKey::Name);
    let descending = sort.map(|sort| sort.descending).unwrap_or(false);

    entries.sort_by(|a, b| {
        // The direction is the column's own and is applied to the column alone: the name
        // beneath it is a tie-break, and a listing that is sorted newest-first still reads
        // its names the same way round. Reversing the whole order instead would turn every
        // pair of equal dates or sizes upside down as well, which is a listing no file
        // manager shows.
        let by_key = match key {
            // The name is the column here rather than the tie-break beneath one, so it
            // takes the direction below like every other column and a listing sorted
            // Z-to-A reads the way it says it does.
            SortKey::Name => natural_cmp(&a.path, &b.path),
            SortKey::DateModified => {
                // A file whose date could not be read has none, and goes last in an ascending
                // walk — the same place a date of the beginning of time would put it, and the
                // place a listing puts a file it has not stamped.
                let left = a.facts.as_ref().and_then(|facts| facts.modified);
                let right = b.facts.as_ref().and_then(|facts| facts.modified);
                left.cmp(&right)
            }
            SortKey::Size => {
                let left = a.facts.as_ref().map(|facts| facts.bytes);
                let right = b.facts.as_ref().map(|facts| facts.bytes);
                left.cmp(&right)
            }
            SortKey::FileType => {
                // The kind's own claim, then the name: a type column groups by the extension
                // Explorer shows, which is the extension, and the kind is as close as this app
                // gets to it without reading a registry.
                kind_rank(a.kind).cmp(&kind_rank(b.kind))
            }
        };

        // The direction belongs to the column and not to the name beneath it, so it is
        // applied to the column's own answer before the two are put together: a listing
        // sorted newest-first still reads its equal dates the same way round. Reversing
        // after the name is in has been decided would turn every pair of equal dates or
        // sizes upside down as well, which is a listing no file manager shows.
        let by_key = if descending {
            by_key.reverse()
        } else {
            by_key
        };

        by_key.then_with(|| natural_cmp(&a.path, &b.path))
    });
}

/// A kind's place in the order the claims are asked in, which is the order a file manager
/// would group them by: a video before a sound, a page before the boxes, the text, the fonts
/// and the pictures last (see `routing::CLAIMS`). `PreviewType` is not `Ord`, and a sort
/// that made it one would put a derived order where a hand wrote one.
fn kind_rank(kind: PreviewType) -> u8 {
    match kind {
        PreviewType::Videos => 0,
        PreviewType::Audio => 1,
        PreviewType::Ebook => 2,
        PreviewType::Archives => 3,
        PreviewType::Peazip => 4,
        PreviewType::Calibre => 5,
        PreviewType::Document => 6,
        PreviewType::Libre => 7,
        PreviewType::Magick => 8,
        PreviewType::Design => 9,
        PreviewType::Vector => 10,
        PreviewType::Text => 11,
        PreviewType::Fonts => 12,
        PreviewType::Images => 13,
    }
}

/// Two paths in the order a file manager shows names, which is not the order their bytes
/// sort in: the runs of digits in a name are compared as numbers, so `img2` comes before
/// `img10` and `part 3` before `part 10`.
///
/// This is the same comparison Explorer's own `Name` column makes, and it is written here
/// rather than linked from `shlwapi` for the reason the rest of this app names its own
/// dependencies: one small testable function is not a dependency, and `StrCmpLogicalW` is
/// not something this build has any other use for.
///
/// The run of digits is compared by value where both are digits, and the runs are trimmed of
/// leading zeros so `img007` and `img7` are one number and not two spellings of it. A name
/// that is all digits is compared as one number, so a folder of numbered screenshots is in
/// order rather than in the order the digits happen to fall.
fn natural_cmp(a: &Path, b: &Path) -> Ordering {
    let left = a.file_name().and_then(|name| name.to_str()).unwrap_or_default();
    let right = b
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or_default();

    natural_str_cmp(left, right)
}

/// The comparison itself, over two names, so it can be tested without two paths.
fn natural_str_cmp(a: &str, b: &str) -> Ordering {
    let mut left = a.chars().peekable();
    let mut right = b.chars().peekable();

    loop {
        match (left.peek().copied(), right.peek().copied()) {
            (None, None) => return Ordering::Equal,
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (Some(l), Some(r)) => {
                let left_digits = l.is_ascii_digit();
                let right_digits = r.is_ascii_digit();

                if left_digits && right_digits {
                    let left_run = take_digits(&mut left);
                    let right_run = take_digits(&mut right);
                    // Trimmed of their leading zeros so the two are compared as the numbers
                    // they spell rather than as their spellings.
                    let left_number = left_run.trim_start_matches('0');
                    let right_number = right_run.trim_start_matches('0');

                    match left_number.len().cmp(&right_number.len()).then_with(|| {
                        left_number.cmp(right_number)
                    }) {
                        Ordering::Equal => continue,
                        other => return other,
                    }
                } else {
                    // A letter and a digit are not compared by their case: an upper-case run
                    // sorts before a lower-case one, which is what a file manager does and
                    // what makes `README` sit above `readme.txt`.
                    let ordering = l
                        .to_ascii_lowercase()
                        .cmp(&r.to_ascii_lowercase())
                        .then_with(|| l.cmp(&r));
                    match ordering {
                        Ordering::Equal => {
                            left.next();
                            right.next();
                        }
                        other => return other,
                    }
                }
            }
        }
    }
}

/// The run of digits at the front of a name, taken off it and handed back as it is written.
fn take_digits(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> String {
    let mut run = String::new();
    while let Some(c) = chars.peek().copied() {
        if !c.is_ascii_digit() {
            break;
        }
        run.push(c);
        chars.next();
    }
    run
}

/// Whether two paths are the same file, told apart only by what a walk compares them on: a
/// verbatim `\\?\` path and a plain one are the same file, and a walk handed the other
/// spelling of the pinned file's path must not decide it is not on its own list — which
/// would send the first `Next` to the top of the folder rather than to the file beside it.
fn same_file(a: &Path, b: &Path) -> bool {
    a == b || plain_path(a).eq_ignore_ascii_case(&plain_path(b))
}

/// A path with the verbatim prefix the Shell canonicalizes to taken off — the same
/// adjustment `webview_preview` makes before it points a browser at a file and
/// `office_render` makes before it hands a document to Office (see `plain_path` there).
fn plain_path(path: &Path) -> String {
    let text = path.to_string_lossy();

    match text.strip_prefix(r"\\?\UNC\") {
        Some(rest) => format!(r"\\{rest}"),
        None => text
            .strip_prefix(r"\\?\")
            .unwrap_or(&text)
            .to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::config::DEFAULT_PIN_NAV_FILE_TYPES;

    /// A folder of a test's own under the temp folder, named by the test and cleared first,
    /// so two tests walking two files cannot see each other's.
    fn folder_named(name: &str) -> PathBuf {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-pin-navigation-tests")
            .join(name);
        let _ = std::fs::remove_dir_all(&folder);
        std::fs::create_dir_all(&folder).expect("a folder a test can write to");
        folder
    }

    /// A file of a given size in a folder, written the way a test writes one. The size is
    /// what a size-ordered walk reads, and it is the one property of a file a test can set
    /// without asking the filesystem to remember something: a date is a timestamp the
    /// filesystem keeps to itself, so the date order is tested where the date lives.
    fn file_of(folder: &Path, name: &str, bytes: usize) -> PathBuf {
        let path = folder.join(name);
        std::fs::write(&path, vec![0u8; bytes]).expect("a file a test can write");
        path
    }

    /// The names of a walk, in the order it walks them.
    fn names(list: &[PathBuf]) -> Vec<String> {
        list.iter()
            .map(|path| {
                path.file_name()
                    .map(|name| name.to_string_lossy().into_owned())
                    .unwrap_or_default()
            })
            .collect()
    }

    /// A walk of a folder, with whatever sort the view of it was in. The walk's own map is
    /// emptied first, so each call reads the folder rather than a folder a test before it
    /// left behind, and a test that cares about the order sets one through `walked_in`.
    fn walked(folder: &Path, current: &Path, mode: PinNavFileTypes) -> Vec<String> {
        walked_in(folder, current, None, mode)
    }

    /// The same walk with a sort in place, which is what the view would have said.
    fn walked_in(
        folder: &Path,
        current: &Path,
        sort: Option<ViewSort>,
        mode: PinNavFileTypes,
    ) -> Vec<String> {
        FOLDERS.lock().expect("the walk's own map").clear();

        let config = AppConfig {
            pin_nav_file_types: mode,
            ..AppConfig::default()
        };

        let entries = scan(folder, &config, sort);
        let mut ordered = entries;
        order(&mut ordered, sort);

        let wanted = routing::kind_of(current, &config).map(routing::nav_category);
        let walked: Vec<PathBuf> = ordered
            .iter()
            .filter(|entry| {
                mode == PinNavFileTypes::All
                    || (wanted.is_some()
                        && wanted == Some(routing::nav_category(entry.kind)))
            })
            .map(|entry| entry.path.clone())
            .collect();

        names(&walked)
    }

    /// One entry of a walk, built by hand so that a date order can be set up without a
    /// filesystem behind it: a date and a size are facts of the entry rather than of a file,
    /// which is what makes a date order testable at all.
    fn entry(name: &str, kind: PreviewType, bytes: u64, days_ago: u64) -> NavEntry {
        NavEntry {
            path: PathBuf::from(name),
            kind,
            facts: Some(Facts {
                modified: Some(SystemTime::UNIX_EPOCH + std::time::Duration::from_secs(86_400 * days_ago)),
                bytes,
            }),
        }
    }

    /// The names of a set of entries in the order a sort put them in.
    fn ordered(entries: Vec<NavEntry>, sort: Option<ViewSort>) -> Vec<String> {
        let mut entries = entries;
        order(&mut entries, sort);
        names(
            &entries
                .iter()
                .map(|entry| entry.path.clone())
                .collect::<Vec<PathBuf>>(),
        )
    }

    /// The pin's own buttons walk a folder in the order a file manager shows names, and a
    /// file manager does not put `img10` before `img2` the way a byte order does. This is
    /// what the whole walk falls back to, and it is what a `Name` column reproduces.
    #[test]
    fn a_walk_orders_numbers_the_way_a_hand_reads_them() {
        assert_eq!(natural_str_cmp("img2.png", "img10.png"), Ordering::Less);
        assert_eq!(natural_str_cmp("img10.png", "img2.png"), Ordering::Greater);
        assert_eq!(natural_str_cmp("img2.png", "img2.png"), Ordering::Equal);

        assert_eq!(natural_str_cmp("part 3.txt", "part 10.txt"), Ordering::Less);
        assert_eq!(natural_str_cmp("2.txt", "10.txt"), Ordering::Less);
        assert_eq!(natural_str_cmp("img007.png", "img7.png"), Ordering::Equal);
        assert_eq!(natural_str_cmp("a1b2.png", "a1b10.png"), Ordering::Less);

        // A name is compared without regard to case first, so a mixed-case run sorts beside
        // the lower-case one rather than after every one of them.
        assert_eq!(natural_str_cmp("README.md", "readme.txt"), Ordering::Less);
        assert_eq!(natural_str_cmp("notes.txt", "notes.txt.bak"), Ordering::Less);
    }

    /// A folder of exported frames, walked the way a file manager walks one: `img2` before
    /// `img10`, where a plain comparison of the same three names — which is what this app's
    /// own `PathBuf` ordering would have given — puts `img10` second.
    #[test]
    fn a_folder_of_numbered_names_is_walked_in_the_order_they_are_read_in() {
        let folder = folder_named("numbered");
        file_of(&folder, "img10.png", 10);
        file_of(&folder, "img2.png", 10);
        file_of(&folder, "img1.png", 10);

        let walked = walked(&folder, &folder.join("img1.png"), PinNavFileTypes::All);
        assert_eq!(walked, vec!["img1.png", "img2.png", "img10.png"]);

        let mut by_byte = vec![
            PathBuf::from("img1.png"),
            PathBuf::from("img2.png"),
            PathBuf::from("img10.png"),
        ];
        by_byte.sort();
        assert_eq!(
            names(&by_byte),
            vec!["img1.png", "img10.png", "img2.png"],
            "which is the order a plain comparison gives, and the one a walk must not have"
        );
    }

    /// The four columns a listing can be in, each one ascending and each one descending, are
    /// the four orders a walk can be in: a walk that reproduced one and not the others would
    /// be a walk that matched the screen a quarter of the time. The entries are built by
    /// hand because a file's date is a timestamp the filesystem keeps to itself, and setting
    /// one is a call this app makes nowhere and would not make for a test either.
    #[test]
    fn every_column_a_listing_is_sorted_by_is_reproduced_both_ways_round() {
        // Written so that every column's answer differs from every other column's: the
        // names, the sizes, the dates and the kinds are in four different orders.
        let entries = || {
            vec![
                entry("clip.mp4", PreviewType::Videos, 300, 30),
                entry("note.txt", PreviewType::Text, 100, 10),
                entry("shot.png", PreviewType::Images, 200, 20),
                entry("atlas.pdf", PreviewType::Ebook, 400, 40),
            ]
        };

        for (key, ascending) in [
            (
                SortKey::Name,
                vec!["atlas.pdf", "clip.mp4", "note.txt", "shot.png"],
            ),
            // The dates run from one day ago to four, so ascending by date is the youngest
            // first and is the same order the ages give — which is the opposite end of the
            // sizes above, and a walk that mixed the two up could not pass both.
            (
                SortKey::DateModified,
                vec!["note.txt", "shot.png", "clip.mp4", "atlas.pdf"],
            ),
            (
                SortKey::Size,
                vec!["note.txt", "shot.png", "clip.mp4", "atlas.pdf"],
            ),
            // A type column groups by kind, and the kinds are grouped in the order the
            // claims are asked in: the video, the page, the text, and the picture last
            // (see `kind_rank`).
            (
                SortKey::FileType,
                vec!["clip.mp4", "atlas.pdf", "note.txt", "shot.png"],
            ),
        ] {
            let up = ordered(entries(), Some(ViewSort { key, descending: false }));
            let down = ordered(entries(), Some(ViewSort { key, descending: true }));

            assert_eq!(up, ascending, "{key:?} ascending");

            // No two of the files above share a value in the column, so the two walks are
            // each other's reverse here. Where two files do share one, the name beneath the
            // column is what tells them apart and it is read the same way round either way —
            // which is what the second half of this test is about.
            let mut reversed = up.clone();
            reversed.reverse();
            assert_eq!(
                down, reversed,
                "{key:?} descending, which is the ascending walk the other way round"
            );
        }

        // The two walks are the same walk only because no two of the files above share a
        // value. A descending size column does not read its names backwards: two files of
        // one length sit in the name order whichever way the column runs, which is what a
        // listing shows and what reversing the whole order would lose.
        let tied = vec![
            entry("b.png", PreviewType::Images, 10, 1),
            entry("a.png", PreviewType::Images, 10, 1),
        ];
        assert_eq!(
            ordered(tied.clone(), Some(ViewSort { key: SortKey::Size, descending: false })),
            vec!["a.png", "b.png"],
            "equal sizes fall to the name, the way round it is read"
        );
        assert_eq!(
            ordered(tied, Some(ViewSort { key: SortKey::Size, descending: true })),
            vec!["a.png", "b.png"],
            "and a descending column does not turn the names about under them"
        );
    }

    /// A view that could name no column, a view whose items were dragged where the user put
    /// them, and a search results view all answer the same way: there is no order to
    /// reproduce, and the name order is what the walk falls back to — which is also what
    /// Explorer shows for a default view, so the two agree where it matters.
    #[test]
    fn a_listing_with_no_order_to_copy_is_walked_in_name_order() {
        let folder = folder_named("no-order");
        file_of(&folder, "img10.png", 10);
        file_of(&folder, "img2.png", 20);
        file_of(&folder, "img1.png", 30);

        // No sort remembered, which is what a folder nothing has been hovered in is: an
        // icon view, a view whose items were moved, a search — and a folder the hook has
        // not been into at all.
        assert_eq!(
            walked(&folder, &folder.join("img1.png"), PinNavFileTypes::All),
            vec!["img1.png", "img2.png", "img10.png"]
        );

        // And the same three names through the sort itself, with nothing to say: which is
        // what a view that named no column is remembered as.
        assert_eq!(
            ordered(
                vec![
                    entry("img10.png", PreviewType::Images, 10, 1),
                    entry("img2.png", PreviewType::Images, 20, 2),
                    entry("img1.png", PreviewType::Images, 30, 3),
                ],
                None
            ),
            vec!["img1.png", "img2.png", "img10.png"],
            "no sort at all is the name order, which is what a default view shows"
        );
    }

    /// `All` walks a folder end to end and `Category` walks the part of it that is the same
    /// kind of thing as the pinned file: under the first, a video steps to the sound beside
    /// it; under the second, it does not — which is the whole of what the switch is for, and
    /// both lists come out of one scan of the folder.
    #[test]
    fn the_two_answers_narrow_the_same_walk_by_different_rules() {
        let folder = folder_named("two-answers");
        file_of(&folder, "a.mp4", 10);
        let song = file_of(&folder, "b.mp3", 10);
        file_of(&folder, "c.mp3", 10);
        file_of(&folder, "d.txt", 10);

        let film = folder.join("a.mp4");
        assert_eq!(
            walked(&folder, &film, PinNavFileTypes::All),
            vec!["a.mp4", "b.mp3", "c.mp3", "d.txt"],
            "`All` walks a folder end to end, so a video sits beside a sound"
        );

        assert_eq!(
            walked(&folder, &song, PinNavFileTypes::Category),
            vec!["b.mp3", "c.mp3"],
            "`Category` walks only what is the same kind of thing, which for a sound is the \
             sounds"
        );
    }

    /// A kind the tray has switched off is not walked, whatever else it is: a file a pin
    /// could step onto and then show nothing of is a step into nowhere, and the walk is of
    /// the files this configuration previews rather than of the files on the disk.
    #[test]
    fn a_kind_the_tray_has_switched_off_is_not_stepped_onto() {
        let folder = folder_named("switched-off");
        file_of(&folder, "a.txt", 10);
        file_of(&folder, "b.mp3", 10);

        FOLDERS.lock().expect("the walk's own map").clear();

        let mut config = AppConfig::default();
        config.audio_preview_enabled = false;
        let current = folder.join("a.txt");

        let list = list_for(&current, &config)
            .map(|list| names(&list))
            .unwrap_or_default();
        assert_eq!(list, vec!["a.txt"], "a switched-off kind is not a step");
    }

    /// A file this build has nothing to draw it with is not a step either: a `.epub` on a
    /// machine with no Calibre is not something a `Next` can land on and show, and a walk
    /// that offered one would step the pin onto a file whose preview is an error. The
    /// `.psd` beside it is a step whether or not any converter is installed, because its
    /// own format keeps a picture of the whole document.
    #[test]
    fn a_file_nothing_on_this_machine_can_be_drawn_with_is_not_stepped_onto() {
        let folder = folder_named("no-reader");
        file_of(&folder, "a.txt", 10);
        file_of(&folder, "b.psd", 10);
        let book = file_of(&folder, "c.epub", 10);
        file_of(&folder, "d.png", 10);

        let list = walked(&folder, &folder.join("a.txt"), PinNavFileTypes::All);

        assert!(
            list.contains(&"b.psd".to_string()),
            "a design document keeps a picture of itself, so it needs no engine: {list:?}"
        );
        assert!(list.contains(&"d.png".to_string()));

        // A book is walked or it is not by what this machine can convert it with, which is a
        // question about the machine and not about the walk — so what is asked here is the
        // rule, not the answer. A test that asked the answer would pass on a machine without
        // Calibre and fail on the author's own.
        assert_eq!(
            list.contains(&"c.epub".to_string()),
            !routing::readers_for(PreviewType::Calibre, &book).is_empty(),
            "a book is a step exactly when an engine on this machine can convert it"
        );

        // And nothing about the walk touches the file on disk: it is a list of names, not a
        // read of what they hold.
        assert!(book.exists());
    }

    /// A folder is a ring: the first file's previous is the last, and the last file's next
    /// is the first. A pin parked at either end of a folder is at the end of it rather than
    /// stuck there, which is the whole of what "wrap around" buys.
    #[test]
    fn a_folder_is_walked_round_and_both_ends_come_back_to_the_other() {
        let folder = folder_named("wrap");
        file_of(&folder, "a.txt", 10);
        let b = file_of(&folder, "b.txt", 10);
        let c = file_of(&folder, "c.txt", 10);
        let a = folder.join("a.txt");

        let list = walked(&folder, &b, PinNavFileTypes::All)
            .into_iter()
            .map(|name| folder.join(name))
            .collect::<Vec<PathBuf>>();

        assert_eq!(step_to(&b, &list, 1).as_deref(), Some(c.as_path()));
        assert_eq!(
            step_to(&c, &list, 1).as_deref(),
            Some(a.as_path()),
            "past the last file is the first"
        );
        assert_eq!(
            step_to(&a, &list, -1).as_deref(),
            Some(c.as_path()),
            "and before the first is the last"
        );
        assert_eq!(step_to(&b, &list, 0).as_deref(), Some(b.as_path()));
        assert_eq!(step_to(&b, &[], 1), None, "an empty walk steps nowhere");
    }

    /// A pinned file the walk does not have — a file whose kind was switched off after the
    /// pin was taken up, or one no kind claims — has no position to step from. The step is
    /// taken from the end it points at: `Next` to the first file, which is where the hand is
    /// looking, and `Previous` to the last, which is the same walk the other way round. The
    /// pinned file is not lost by any of it.
    ///
    /// The walk is the folder's own rather than the pinned file's category, because under
    /// `Category` a file that is on the walk is on it by being of the same kind — a walk a
    /// pinned file is missing from is one its own kind cannot produce, and the case that does
    /// produce it is a kind the tray has switched off since.
    #[test]
    fn a_file_the_walk_does_not_have_steps_on_from_the_end_it_points_at() {
        let folder = folder_named("off-list");
        let a = file_of(&folder, "a.mp3", 10);
        let b = file_of(&folder, "b.mp3", 10);
        let c = file_of(&folder, "c.mp3", 10);
        // A fourth file of a kind this app previews and this build has no reader for, so it
        // is not on the walk at all — the only way a file that *can* be previewed is not on it.
        let unreadable = file_of(&folder, "d.cdr", 10);

        let list = vec![a.clone(), b.clone(), c.clone()];

        assert!(
            !list.contains(&folder.join("d.cdr")),
            "the file the pin is on is a kind nothing on this machine can draw"
        );
        assert_eq!(
            step_to(&unreadable, &list, 1).as_deref(),
            Some(a.as_path()),
            "`Next` from off the list is the first file"
        );
        assert_eq!(
            step_to(&unreadable, &list, -1).as_deref(),
            Some(c.as_path()),
            "and `Previous` from off the list is the last"
        );
        assert!(
            unreadable.exists(),
            "and the file pinned is still there to be pinned back to"
        );
    }

    /// A verbatim path and the plain spelling of it are one file, and a walk handed one of
    /// them must find the other on its own list: the Shell hands this app its paths in the
    /// verbatim form, and a walk that did not recognise one would decide the pinned file is
    /// not on the list and send the first `Next` to the top of the folder.
    #[test]
    fn a_verbatim_path_is_the_same_file_as_the_plain_spelling_of_it() {
        let plain = Path::new(r"C:\art\clip.mp4");
        let verbatim = Path::new(r"\\?\C:\art\clip.mp4");
        assert!(same_file(plain, verbatim));
        assert!(same_file(verbatim, plain));
        assert!(!same_file(plain, Path::new(r"C:\art\other.mp4")));

        assert!(same_file(
            Path::new(r"\\?\UNC\server\share\clip.mp4"),
            Path::new(r"\\server\share\clip.mp4")
        ));

        // And the step is taken from the file's own place rather than from the end. The list
        // holds the paths a walk of the folder produces, which are full paths in the folder's
        // own spelling — so this is the case the app meets: a pin whose path came back from
        // the Shell in the verbatim form, walking a list of plain ones.
        let list = vec![
            PathBuf::from(r"C:\art\a.mp3"),
            PathBuf::from(r"C:\art\b.mp3"),
        ];
        assert_eq!(
            step_to(Path::new(r"\\?\C:\art\b.mp3"), &list, -1).as_deref(),
            Some(Path::new(r"C:\art\a.mp3")),
            "the verbatim spelling of the second file steps back to the first, not to the end"
        );
        assert_eq!(
            step_to(Path::new(r"\\?\C:\art\a.mp3"), &list, 1).as_deref(),
            Some(Path::new(r"C:\art\b.mp3")),
            "and forwards the other way round"
        );
    }

    /// The default is `All`, so a pin walks a folder end to end where the configuration says
    /// nothing at all — which is what a `config.ini` written before this setting existed is
    /// read as, and what a file naming a value this build cannot read is left at.
    #[test]
    fn a_pin_walks_everything_unless_it_is_told_otherwise() {
        assert_eq!(DEFAULT_PIN_NAV_FILE_TYPES, PinNavFileTypes::All);
        assert_eq!(PinNavFileTypes::All.as_str(), "all");
        assert_eq!(PinNavFileTypes::Category.as_str(), "category");
    }
}

