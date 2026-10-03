//! What an entry of a list may be, and the one sanitiser the four rules below are written in.
//!
//! A row names the rule rather than a function of its own, so what is here is the rule and the
//! characters it is read by — [`Entries`], the four tables, and the function they share — and
//! nothing else: no list, no section, and no knowledge of what any format is called.

/// What the entries of one list are allowed to be, which is the whole of what sixteen
/// sanitizers differed on once fourteen of them were one function.
///
/// A choice of characters rather than of function on purpose: a row that named a sanitiser could
/// name one that does not exist, and the difference between the four was never a different
/// function — it was a different set of characters to read an entry by.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Entries {
    /// A bare extension: its alphanumerics and `+`, `-` and `_`. Twelve of the sixteen lists are
    /// this and nothing else, and every one of them reads the same way.
    Bare,
    /// A bare extension, or a compound of two of them. The archive list alone, because `tar.gz`
    /// is a name rather than an extension — the last dot of `sources.tar.gz` is `gz`, which is
    /// not an archive on its own — so a sanitiser that dropped the dot would silently stop the
    /// list claiming the format it exists to claim.
    Compound,
    /// A bare extension, or one holding a `#`. The text list alone, and only so that a C# or F#
    /// project is a text file rather than an extension nobody can type.
    Hashed,
    /// A file name rather than an extension: a dot begins the name instead of changing it, so
    /// `.gitignore` has no extension at all and `cmakelists.txt` has one in the middle of it.
    Name,
}

impl Entries {
    /// One list as the entries a lookup compares against, by the rule the row names.
    ///
    /// Every list in the running app is read through here and nowhere else — the four rules
    /// below are the four this dispatches to, under the names the rest of the tree has always
    /// called them by.
    pub(super) fn sanitize(self, list: &str) -> Vec<String> {
        match self {
            Entries::Bare => sanitize_extension_list(list),
            Entries::Compound => sanitize_archive_extension_list(list),
            Entries::Hashed => sanitize_extensions(list),
            Entries::Name => sanitize_names(list),
        }
    }
}

/// The characters an ordinary extension list admits besides its alphanumerics.
///
/// Deliberately without `#` and without `.`, which are the two rules that need them: `#` for a
/// name like `C#`, `.` for a compound like `tar.gz`.
const PLAIN_EXTENSION_CHARS: &[char] = &['+', '-', '_'];
/// The archive list's own set, which is the one list that may hold a dotted compound name.
const COMPOUND_EXTENSION_CHARS: &[char] = &['.', '+', '-', '_', '#'];
/// The text list's own set, which is the one extension list that may hold a `#`.
const HASHED_EXTENSION_CHARS: &[char] = &['+', '-', '_', '#'];
/// And the name list's, for entries that are file names rather than extensions.
const NAME_CHARS: &[char] = &['.', '-', '_'];

/// One list split into the entries a lookup compares against, each read by the rule `allowed`
/// names.
///
/// A leading dot is accepted because `py` and `.py` are both what a user might type, and an entry
/// that is not a plausible name or extension is dropped so a stray path or sentence in the list
/// cannot turn into a match — which is what an entry is read as everywhere else in this app.
fn sanitized_list(list: &str, allowed: &[char]) -> Vec<String> {
    let mut entries: Vec<String> = Vec::new();

    for entry in list.split(',') {
        let entry = entry.trim().trim_start_matches('.').to_lowercase();
        if entry.is_empty()
            || !entry
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || allowed.contains(&c))
        {
            continue;
        }

        // A list is a set, and a user who typed a name twice gets it once: the second entry would
        // be a second match where one file already satisfies the first.
        if !entries.iter().any(|held| held == &entry) {
            entries.push(entry);
        }
    }

    entries
}

/// The sanitiser for the twelve lists that are a bare extension and nothing else.
pub fn sanitize_extension_list(list: &str) -> Vec<String> {
    sanitized_list(list, PLAIN_EXTENSION_CHARS)
}

/// The archive list's own sanitiser, which is the one list that holds a dotted compound name.
///
/// `tar.gz` is a name rather than an extension, and `archive_formats::matches_archive_list`
/// matches it against the end of a whole file name — so a sanitiser that dropped the dot would
/// silently stop the list claiming the format it exists to claim.
pub fn sanitize_archive_extension_list(list: &str) -> Vec<String> {
    sanitized_list(list, COMPOUND_EXTENSION_CHARS)
}

/// The text list's own sanitiser, which is the one extension list that admits a `#`, for `C#`.
pub fn sanitize_extensions(list: &str) -> Vec<String> {
    sanitized_list(list, HASHED_EXTENSION_CHARS)
}

/// The name list's own sanitiser, for a list whose entries are file names and not extensions.
pub fn sanitize_names(list: &str) -> Vec<String> {
    sanitized_list(list, NAME_CHARS)
}
