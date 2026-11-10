use super::*;

/// The folder these tests keep their files in, made fresh: the app's own scratch under `%TEMP%`,
/// never a folder on the machine that a run would be reading (see ARCHITECTURE.md, "Test isolation").
pub(super) fn a_folder_of_our_own(name: &str) -> PathBuf {
    let folder = std::env::temp_dir().join(format!(
        "rust-hover-preview-everything-{}-{name}",
        std::process::id()
    ));
    let _ = std::fs::remove_dir_all(&folder);
    std::fs::create_dir_all(&folder).expect("a folder of the test's own");
    folder
}

/// A file whose own first bytes say it is a picture, which is what every route here gates on: the
/// name alone is not the kind (see `crate::formats::head::Facts`).
pub(super) fn a_picture(folder: &Path, name: &str) -> PathBuf {
    let path = folder.join(name);
    std::fs::write(&path, b"\x89PNG\x0d\x0a\x1a\x0a").expect("a picture");
    path
}

/// The cells a details row of a search over one folder is drawn with: the name, the folder the file
/// is in, and the three columns Everything's default layout puts beside them.
fn a_row(folder: &Path, name: &str) -> Vec<String> {
    vec![
        name.to_string(),
        folder.to_string_lossy().into_owned(),
        "35,131 KB".to_string(),
        "7/5/2026 9:25 PM".to_string(),
        "7/5/2026 9:25 PM".to_string(),
    ]
}

/// **The row's own text is the whole witness**, and the two facts it holds are the name and the folder
/// — a search spans folders, so nothing else in the listing says where the file is.
#[test]
fn the_name_and_the_folder_cells_name_the_file() {
    let folder = a_folder_of_our_own("name-and-folder");
    let picture = a_picture(&folder, "holiday.png");

    assert_eq!(
        row_file(&a_row(&folder, "holiday.png")),
        normalize_media_path(picture),
        "a name and the folder it is in is the file the row stands for"
    );
}

/// A row drawn for a folder is a name with nothing to preview: a folder is no file's name, and a
/// preview of whatever happens to be inside it was never pointed at.
#[test]
fn a_folder_row_answers_nothing() {
    let folder = a_folder_of_our_own("folder-row");
    let inside = folder.join("album");
    std::fs::create_dir_all(&inside).expect("a folder of the test's own");
    a_picture(&inside, "holiday.png");

    assert_eq!(
        row_file(&a_row(&folder, "album")),
        None,
        "a folder is not a file this app previews"
    );
}

/// **A size or a date is never read as the folder**, whatever the user's columns are: the folder is the
/// first cell that *is* one — absolute, and a directory — and the details columns Everything draws by
/// default are neither. Nothing is mistaken for a path, and a row whose folder cell has been removed
/// answers nothing at all rather than joining a name onto a number.
#[test]
fn a_size_or_a_date_is_never_taken_for_the_folder() {
    let folder = a_folder_of_our_own("no-folder-cell");
    a_picture(&folder, "holiday.png");

    for cells in [
        vec![
            "holiday.png".to_string(),
            "35,131 KB".to_string(),
            "7/5/2026 9:25 PM".to_string(),
        ],
        vec!["holiday.png".to_string()],
        vec!["holiday.png".to_string(), "".to_string()],
        // A relative cell is not a place: it is read against whatever folder this process happens to
        // be in, which is nobody's answer.
        vec!["holiday.png".to_string(), ".".to_string()],
    ] {
        assert_eq!(
            row_file(&cells),
            None,
            "{cells:?} names no folder, so it names no file"
        );
    }
}

/// The folder need not be the second cell: the columns of a list are the user's arrangement, so the
/// path is found wherever it is drawn.
#[test]
fn the_folder_is_found_wherever_the_columns_put_it() {
    let folder = a_folder_of_our_own("folder-not-second");
    let picture = a_picture(&folder, "holiday.png");

    let cells = vec![
        "holiday.png".to_string(),
        "35,131 KB".to_string(),
        folder.to_string_lossy().into_owned(),
        "7/5/2026 9:25 PM".to_string(),
    ];
    assert_eq!(row_file(&cells), normalize_media_path(picture));
}

/// A name is a name, whatever it is spelled in: Everything draws the file's own name and the join is
/// the filesystem's.
#[test]
fn a_unicode_name_resolves() {
    let folder = a_folder_of_our_own("unicode");
    let picture = a_picture(&folder, "naïve — 日本語 (.png");

    assert_eq!(
        row_file(&a_row(&folder, "naïve — 日本語 (.png")),
        normalize_media_path(picture),
        "the name is taken as the list draws it, and nothing is escaped or decoded"
    );
}

/// A file whose kind nothing previews answers nothing, and so does a name no file has: the gate is the
/// file's own entry, not the row's word for it.
#[test]
fn a_row_naming_no_previewable_file_answers_nothing() {
    let folder = a_folder_of_our_own("nothing-to-preview");
    std::fs::write(folder.join("notes.txt"), b"a note").expect("a file of the test's own");

    assert_eq!(
        row_file(&a_row(&folder, "gone.png")),
        None,
        "a name with no file behind it is no file"
    );
}

/// **Everything's hit test names a row whether or not the point is on one**, and its answer is only
/// taken where the row's own box holds the point — the two phantom answers measured on the live window
/// are a point above the first row (`iItem` 0) and one below the last row drawn.
#[test]
fn only_a_row_whose_own_box_holds_the_point_answers_it() {
    // Row n's box, drawn the way the live list draws them: 19 px tall, the first one 26 px down.
    let box_of_row = |index: i32| {
        Some(RECT {
            left: 0,
            top: 26 + index * 19,
            right: 1961,
            bottom: 45 + index * 19,
        })
    };

    // The header strip: the hit test answers row 0 for a point above it, and row 0's own box does not
    // reach there.
    assert_eq!(
        row_held_by(POINT { x: 5, y: 10 }, 0, 0, 100, box_of_row),
        None,
        "a point above the first row is on no row, whatever the hit test names"
    );

    // A point inside a row, on a list that has not been scrolled.
    assert_eq!(
        row_held_by(POINT { x: 5, y: 54 }, 1, 0, 100, box_of_row),
        Some(1),
        "the row the hit test named, where its box holds the point"
    );

    // The same point on a list scrolled by ten rows, where the hit test numbered its answer from the
    // row it has scrolled to: the box of row 1 is drawn above the window, and the row the point is on
    // is the first row drawn plus the answer.
    let scrolled = |index: i32| {
        Some(RECT {
            left: 0,
            top: 26 + (index - 10) * 19,
            right: 1961,
            bottom: 45 + (index - 10) * 19,
        })
    };
    assert_eq!(
        row_held_by(POINT { x: 5, y: 54 }, 1, 10, 100, scrolled),
        Some(11),
        "an answer numbered from the top of a scrolled list is the row below it"
    );

    // A point on a wide row, which Everything answers with `LVHT_NOWHERE` and the right item.
    assert_eq!(
        row_held_by(POINT { x: 1900, y: 54 }, 1, 0, 100, box_of_row),
        Some(1),
        "the flags say nothing; the row's own box is what answers"
    );
}

/// An index a list answered is only a row where the list has one: a results list that emptied under the
/// answer, a hit test's `-1`, and an index past the end are all no row at all.
#[test]
fn an_index_outside_the_list_is_no_row() {
    assert!(!row_is_in_the_list(-1, 10), "no item is no index");
    assert!(!row_is_in_the_list(10, 10), "the row after the last one");
    assert!(row_is_in_the_list(0, 10) && row_is_in_the_list(9, 10));

    assert_eq!(
        hit_rows(-1, 0, 10),
        [None, None],
        "nothing the hit test named is inside the list"
    );
    assert_eq!(
        hit_rows(0, 0, 0),
        [None, None],
        "and nothing is, in a list with no rows"
    );
    assert_eq!(
        row_held_by(POINT { x: 5, y: 30 }, 0, 0, 0, |_| Some(RECT::default())),
        None,
        "a list with no rows holds no row, however the box reads"
    );
}
