//! How a file's name is read, which is the question every list in the app is asked through.
//!
//! The lists themselves — what each of them holds, what this app held of them before, and the
//! four rules their entries are read by — are rows of `crate::formats::lists`, the one table every
//! kind's list is a row of. What is kept here is the part of the question that is not a list: a
//! name is read the same way for every kind, and it is read in two ways because a file's name is
//! not always an extension.
//!
//! A dot file is the reason `lookup_extension` is not one call to `Path::extension`: `.gitignore`
//! has no extension — the dot begins its name — so a leading dot is stripped and what follows is
//! the name the lists are written with, which is what makes `gitignore` in the extension list mean
//! `.gitignore`. And a repository is recognized by files with no extension at all, which is the
//! second list rather than a first one: `lookup_name` reads a whole file name, and the text gate
//! asks both (`matches_text_lists`).
//!
//! The two rules that are not the shared one are the two that are not about extensions: the text
//! extension list admits a `#` for `C#`, and the name list admits a dot inside a name. Both live in
//! `formats::lists` with the lists they read, because a rule for reading an entry and a list of
//! entries to read are one thing here rather than two that have to agree.

use std::path::Path;

/// The four rules a list's entries are read by, under the names the tree has always called them
/// by.
///
/// Nothing in the running app asks for one of them any more — every list is read through its row
/// in `formats::lists`, which is what a list is for — but a test that wants to know what a
/// hand-typed list is read as asks for the rule by name rather than matching on an enum, and
/// there are a good many of those.
#[cfg(test)]
pub use crate::formats::lists::{sanitize_archive_extension_list, sanitize_extension_list};

/// Whether `path` is a page of HTML — the two names a web page goes by, and no other.
///
/// The question is the name alone, exactly as `svg_preview::is_svg_file` asks it: what the
/// browser would be handed is decided by what the file is called, and the switch over it
/// belongs to the configuration rather than to this module (see `render_html`). The
/// extension is read through `lookup_extension`, so `.HTML` and `.html` are one name and a
/// dot file is read the way every other list reads one.
pub fn is_html_extension(path: &Path) -> bool {
    matches!(
        lookup_extension(path).as_deref(),
        Some("htm") | Some("html")
    )
}

/// The extension `path` will be looked up by.
///
/// A dot file is the reason this is not one call to `Path::extension`: `.gitignore`
/// has no extension — the dot begins its name — so a leading dot is stripped and
/// what follows is the name the lists are written with, which is what makes
/// `gitignore` in the extension list mean `.gitignore`.
///
/// It is shared rather than copied for the one other caller that has to agree with a list about
/// what a file is called: the tool a name is routed to inside an installed PeaZip (see
/// `peazip_formats::Backend::of`), which has to read a name the same way the list that claims it
/// did.
pub(crate) fn lookup_extension(path: &Path) -> Option<String> {
    if let Some(extension) = path.extension().and_then(|ext| ext.to_str()) {
        return Some(extension.to_lowercase());
    }

    let name = path.file_name()?.to_str()?;
    let stripped = name.strip_prefix('.')?;
    (!stripped.is_empty() && !stripped.contains('.')).then(|| stripped.to_lowercase())
}

/// The name `path` will be looked up by: its file name in lowercase, with a
/// leading dot dropped, so `.gitignore` and `gitignore` are one entry.
fn lookup_name(path: &Path) -> Option<String> {
    let name = path.file_name()?.to_str()?.to_lowercase();
    if name.is_empty() {
        return None;
    }
    Some(name.trim_start_matches('.').to_string())
}

/// Whether `path` carries an extension the configuration previews as text.
pub fn matches_configured_extension(path: &Path, extensions: &[String]) -> bool {
    let Some(extension) = lookup_extension(path) else {
        return false;
    };

    extensions.contains(&extension)
}

/// Whether `path` carries a name the configuration previews as text.
pub fn matches_configured_name(path: &Path, names: &[String]) -> bool {
    let Some(name) = lookup_name(path) else {
        return false;
    };

    names.contains(&name)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The one rule every list in the app is read through, and it used to be written out
    /// fourteen times: a leading dot is what a user types, an entry that is not a bare
    /// extension is dropped rather than matched against, and a repeat is not a second entry.
    ///
    /// It is pinned here and not only through the row that uses it because the rule is what
    /// every list has in common: the tests beside the ones that do use the rule through a row
    /// pin one hand-typed string per list, and a rule cannot be pinned by a sample of the
    /// strings it reads.
    #[test]
    fn a_sanitized_list_drops_a_leading_dot_a_duplicate_and_anything_that_is_not_an_extension() {
        let typed = sanitize_extension_list(" .ZIP , zip,,book*.azw,epub,..,tar.gz");

        // `book*.azw` is a path fragment and `tar.gz` a compound name: neither is a bare
        // extension, so neither is in a list that is matched against one.
        assert_eq!(typed, vec!["zip", "epub"]);
    }

    /// `tar.gz` is a name rather than an extension, and the archive list matches it against
    /// the end of a whole file name - so a sanitiser that dropped the dot would silently
    /// stop the list claiming the format it exists to claim. This is the only difference
    /// between the two lists, and it is the reason they are two functions.
    #[test]
    fn the_archive_list_is_the_one_that_keeps_a_dotted_compound_name() {
        let typed = sanitize_archive_extension_list(" .ZIP , zip,,nonsense*,docx,tar.gz");

        assert_eq!(typed, vec!["zip", "docx", "tar.gz"]);
    }

    /// A page of HTML is a page of HTML under either of its two names and under no other
    /// one: `.xhtml` is XML, which the text preview still reads, and a file with no
    /// extension at all is a name rather than an extension.
    #[test]
    fn a_page_of_html_is_one_of_two_names() {
        for name in [
            "page.html",
            "page.htm",
            "PAGE.HTML",
            "Page.HtM",
            "some/deeper/page.HTML",
        ] {
            assert!(is_html_extension(Path::new(name)), "`{name}` is a page");
        }

        for name in [
            "page.xhtml",
            "page.xml",
            "page.html.gz",
            "page.html5",
            "html",
            "page.txt",
            "Makefile",
        ] {
            assert!(
                !is_html_extension(Path::new(name)),
                "`{name}` is not one of the two names"
            );
        }
    }
}
