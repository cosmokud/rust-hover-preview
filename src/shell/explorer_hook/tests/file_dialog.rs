use super::*;

/// The folder these tests keep their files in, made fresh: the app's own scratch under `%TEMP%`,
/// never a folder on the machine that a run would be reading (see ARCHITECTURE.md, "Test isolation").
fn a_folder_of_our_own(name: &str) -> PathBuf {
    let folder = std::env::temp_dir().join(format!(
        "rust-hover-preview-dialog-{}-{name}",
        std::process::id()
    ));
    let _ = std::fs::remove_dir_all(&folder);
    std::fs::create_dir_all(&folder).expect("a folder of the test's own");
    folder
}

/// A file whose own first bytes say it is a picture, which is what every route here gates on:
/// the name alone is not the kind (see `crate::formats::head::Facts`).
fn a_picture(folder: &Path, name: &str) -> PathBuf {
    let path = folder.join(name);
    std::fs::write(&path, b"\x89PNG\x0d\x0a\x1a\x0a").expect("a picture");
    path
}

/// **The address prefix is the whole of what makes a breadcrumb an address.** `1001` names the
/// address toolbar in an `Open`, and the name field in a `Save As` — so a string that happens to
/// be a path is not the dialog's folder, and reading it as one would answer a hover from whatever
/// the user had typed into a name field.
#[test]
fn an_address_is_only_an_address_with_its_prefix() {
    assert_eq!(
        address_from_breadcrumb("Address: X:\\dir"),
        Some(PathBuf::from("X:\\dir")),
        "the address toolbar's own name is its folder"
    );
    assert_eq!(
        address_from_breadcrumb("  address:  X:\\dir  "),
        Some(PathBuf::from("X:\\dir")),
        "the prefix is the System's own label, and its case and spacing are not the place"
    );

    for not_an_address in [
        "X:\\dir",
        "Address:",
        "Address: nothing at all",
        "Folder: X:\\dir",
        "Name: X:\\dir",
        "",
    ] {
        assert_eq!(
            address_from_breadcrumb(not_an_address),
            None,
            "{not_an_address:?} is not an address, and is not stripped into one"
        );
    }
}

/// A name that is a whole path is used as it is: an item that carries its own path — the way a
/// search result's accessible value does — needs no folder to be resolved against, and the dialog
/// is never asked for one.
#[test]
fn a_name_that_is_a_whole_path_is_the_file_it_names() {
    let folder = a_folder_of_our_own("rooted");
    let picture = a_picture(&folder, "a-picture.png");
    let picture = std::fs::canonicalize(&picture).expect("the file the test made");

    // A frame of no window: the dialog answers no folder at all, so what this proves is that the
    // name was answered without one.
    let window = ItemWindow {
        frame: 0,
        drawn_in: 0,
        dialog: true,
    };
    let item = a_dialog_item(&picture.to_string_lossy());

    assert_eq!(
        dialog_item_path(&mut ItemResolver::default(), &window, &item),
        Some(picture),
        "a name that is a whole path is the file it names"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// **The name the dialog shows is joined onto the folder it is showing, and the file's own entry is
/// what vouches for it.** A name that no file answers — a deleted row, a truncated name — is left
/// unanswered rather than guessed at, and so is a file no kind of this build previews.
#[test]
fn a_name_joined_onto_the_folder_is_the_file_only_where_one_is_there() {
    let folder = a_folder_of_our_own("joined");
    let picture = a_picture(&folder, "a-picture.png");
    let picture = std::fs::canonicalize(&picture).expect("the file the test made");
    std::fs::write(folder.join("notes.bin"), b"\x00\x01\x02nothing names this").expect("a file");

    assert_eq!(
        item_path_in_folder(&folder, "a-picture.png"),
        Some(picture),
        "the item's display name inside the folder the dialog shows"
    );
    assert_eq!(
        item_path_in_folder(&folder, "not-there.png"),
        None,
        "a name nothing answers is no file"
    );
    assert_eq!(
        item_path_in_folder(&folder, "notes.bin"),
        None,
        "and a file no kind previews is no preview either"
    );

    // A folder drawn in the list is a name the dialog shows and nothing to show: its row is not a
    // file's, and it is not answered as one.
    std::fs::create_dir_all(folder.join("a-folder")).expect("a folder");
    assert_eq!(item_path_in_folder(&folder, "a-folder"), None);

    let _ = std::fs::remove_dir_all(&folder);
}

/// The dialog hides extensions where the machine does, so an item shown as `holiday` stands for the
/// one media file of that stem — and for nothing at all where the folder holds two, because a stem
/// shared by two files names neither of them.
#[test]
fn a_name_with_no_extension_is_the_one_file_of_that_stem() {
    let folder = a_folder_of_our_own("stem");
    let only = a_picture(&folder, "holiday.png");
    let only = std::fs::canonicalize(&only).expect("the file the test made");

    assert_eq!(
        item_path_in_folder(&folder, "holiday"),
        Some(only.clone()),
        "one media file of the stem is the file the row stands for"
    );
    assert_eq!(
        item_path_in_folder(&folder, "HOLIDAY"),
        Some(only),
        "and the name is matched the way Windows reads one"
    );
    assert_eq!(
        item_path_in_folder(&folder, "not-a-stem"),
        None,
        "a stem no file has is no file"
    );

    // A second media file of the same stem: the name belongs to neither, so nothing is answered.
    std::fs::write(folder.join("holiday.gif"), b"GIF89a\x01\x00\x01\x00")
        .expect("a second picture");
    assert_eq!(
        item_path_in_folder(&folder, "holiday"),
        None,
        "two candidates is a name that names no one file"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// A dialog's item, as `uia_item_from_point` reports one: the display name and the box it is drawn
/// in, and no position at all — which is the fact this whole route exists for.
fn a_dialog_item(name: &str) -> HoveredItem {
    HoveredItem {
        index: None,
        name: name.to_string(),
        value: None,
        bounds: RECT {
            left: 0,
            top: 0,
            right: 200,
            bottom: 20,
        },
        text: None,
        native_window: 0,
    }
}
