use super::everything::{a_folder_of_our_own, a_picture};
use super::*;

/// The cells one row of a Q-Dir pane is drawn with: the name, then the date and the kind its default
/// columns put beside it — no folder, because the pane shows one folder.
fn a_row(name: &str) -> Vec<String> {
    vec![
        name.to_string(),
        "7/5/2026 9:25 PM".to_string(),
        "PNG File".to_string(),
    ]
}

/// **The pane's folder is the whole of what a row is joined onto**, and the file's own entry is what
/// vouches for it — the name alone is a word the pane drew.
#[test]
fn a_name_joined_onto_the_panes_folder_is_the_file() {
    let folder = a_folder_of_our_own("q-dir-joined");
    let picture = a_picture(&folder, "holiday.png");

    assert_eq!(
        row_file_in_folder(&a_row("holiday.png"), &folder),
        normalize_media_path(picture),
        "a name and the folder the pane is showing is the file the row stands for"
    );
}

/// A folder drawn in the pane is a row with nothing to preview, and a name no file answers is no file:
/// `item_path_in_folder` refuses both.
#[test]
fn a_folder_row_and_a_name_no_file_has_answer_nothing() {
    let folder = a_folder_of_our_own("q-dir-refused");
    let inside = folder.join("album");
    std::fs::create_dir_all(&inside).expect("a folder of the test's own");
    a_picture(&inside, "holiday.png");

    assert_eq!(
        row_file_in_folder(&a_row("album"), &folder),
        None,
        "a folder is not a file this app previews"
    );
    assert_eq!(
        row_file_in_folder(&a_row("gone.png"), &folder),
        None,
        "a name with no file behind it is no file"
    );
}

/// **An extensionless name is the shell's hidden-extension listing**, answered by the one media file of
/// that stem — and refused where two files share it, because a name that belongs to two files belongs
/// to neither.
#[test]
fn an_extensionless_name_answers_through_the_one_file_of_that_stem() {
    let folder = a_folder_of_our_own("q-dir-stem");
    let picture = a_picture(&folder, "holiday.png");

    assert_eq!(
        row_file_in_folder(&a_row("holiday"), &folder),
        normalize_media_path(picture),
        "a name the pane draws without an extension is the sole file of that stem"
    );

    a_picture(&folder, "holiday.jpg");
    assert_eq!(
        row_file_in_folder(&a_row("holiday"), &folder),
        None,
        "a second file of one stem is a name that belongs to neither"
    );
}

/// A cell 0 that is a whole path is taken as the path it is, the way a row a shell hand-off names
/// is: the pane's folder is not joined onto it.
#[test]
fn a_cell_that_is_a_whole_path_is_the_file_it_names() {
    let folder = a_folder_of_our_own("q-dir-rooted");
    let picture = a_picture(&folder, "a-picture.png");
    let picture = std::fs::canonicalize(&picture).expect("the file the test made");

    // A folder that is nothing of this file's: the answer cannot have come from the join.
    let other = a_folder_of_our_own("q-dir-rooted-elsewhere");
    assert_eq!(
        row_file_in_folder(&a_row(&picture.to_string_lossy()), &other),
        Some(picture.clone()),
        "an absolute name is the path it names, whatever folder the pane shows"
    );
}

/// A name is a name, whatever it is spelled in: Q-Dir draws the shell's display name and the join is
/// the filesystem's.
#[test]
fn a_unicode_name_resolves() {
    let folder = a_folder_of_our_own("q-dir-unicode");
    let picture = a_picture(&folder, "naïve — 日本語 (.png");

    assert_eq!(
        row_file_in_folder(&a_row("naïve — 日本語 (.png"), &folder),
        normalize_media_path(picture),
        "the name is taken as the pane draws it, and nothing is escaped or decoded"
    );
}

/// **A pane's address bar names a folder, or it names nothing.** It is trimmed and it is absolute and it
/// is there; a placeholder, a URL, a bare relative name and an empty field are none of those, and a path
/// on no drive of this machine is not one either.
#[test]
fn only_an_absolute_folder_that_is_there_is_the_panes_folder() {
    let folder = a_folder_of_our_own("q-dir-address");

    assert_eq!(
        pane_folder_from_edit_text(&folder.to_string_lossy()),
        Some(folder.clone()),
        "an absolute folder is the folder the pane is showing"
    );
    assert_eq!(
        pane_folder_from_edit_text(&format!("  {}  ", folder.to_string_lossy())),
        Some(folder.clone()),
        "and its whitespace is the field's, not the path's"
    );
    assert_eq!(
        pane_folder_from_edit_text(&format!("{}\\", folder.to_string_lossy())),
        Some(folder.clone()),
        "a trailing separator is where the shell puts one, and the folder is the same folder"
    );

    for not_a_folder in [
        &*folder.join("nothing-here").to_string_lossy(),
        "&Abbrechen",
        "",
        "   ",
        "https://example.test/a",
        "holiday.png",
        "C:",
    ] {
        assert_eq!(
            pane_folder_from_edit_text(not_a_folder),
            None,
            "{not_a_folder:?} is not a folder a row could be joined onto"
        );
    }

    let _ = std::fs::remove_dir_all(&folder);
}
