use super::*;

/// The four columns a folder is ordered by that the pin's own buttons can walk, told
/// apart by the system property key rather than by the name a locale would print for
/// them: a view in another language names the same keys, and a column called `Name` and
/// one called `Date created` are the same pair of questions either way.
#[test]
fn a_sort_column_is_told_apart_by_the_property_it_is_on() {
    for (key, expected) in [
        (PKEY_ITEM_NAME_DISPLAY, SortKey::Name),
        (PKEY_DATE_MODIFIED, SortKey::DateModified),
        (PKEY_SIZE, SortKey::Size),
        (PKEY_FILE_TYPE, SortKey::FileType),
    ] {
        assert_eq!(sort_key_of(&key), Some(expected));
    }

    // A column this app has no comparison for is a column it will not half-reproduce, so
    // a view sorted by one falls back to name order rather than to an order that is only
    // sometimes the listing's.
    assert_eq!(
        sort_key_of(&PROPERTYKEY {
            fmtid: GUID::from_u128(0xb725f130_47ef_101a_a5f1_02608c9eebac),
            pid: 15,
        }),
        None,
        "`Date created` is a column this app has no comparison for"
    );
}

/// A view with no sort columns is a normal answer rather than a failure: there is nothing to
/// reproduce there, and the walk falls back to name order.
#[test]
fn a_view_with_no_sort_columns_is_no_sort() {
    assert_eq!(
        sort_from_columns(
            0,
            FWF_AUTOARRANGE.0 as u32,
            Some((PKEY_ITEM_NAME_DISPLAY, SORT_ASCENDING))
        ),
        None,
        "no columns is a view whose order there is nothing to reproduce"
    );
}

/// A search results view is not sorted by a column at all: it is ordered by how relevant
/// each result is to what was typed, across whatever folders the query reached. So the
/// folder a sort would be remembered against is not a folder a search view has, and
/// nothing is remembered for one.
#[test]
fn a_search_view_has_no_folder_to_remember_a_sort_against() {
    assert!(
        is_search_ms_url("search-ms:query=x&crumb=location:C:\\art"),
        "which is what a search view's own URL looks like"
    );
    assert!(
        !is_search_ms_url("file:///C:/art"),
        "and an ordinary folder is not one"
    );
}

/// The direction is part of the sort rather than something read after it: a folder
/// sorted by date with the newest first is a different walk from one with the oldest
/// first, and the buttons move the same either way.
#[test]
fn a_sort_says_which_way_it_runs() {
    let ascending = sort_from_columns(
        1,
        FWF_AUTOARRANGE.0 as u32,
        Some((PKEY_SIZE, SORT_ASCENDING)),
    );
    let descending = sort_from_columns(
        1,
        FWF_AUTOARRANGE.0 as u32,
        Some((PKEY_SIZE, SORT_DESCENDING)),
    );

    assert_eq!(
        ascending,
        Some(ViewSort {
            key: SortKey::Size,
            descending: false
        })
    );
    assert_eq!(
        descending,
        Some(ViewSort {
            key: SortKey::Size,
            descending: true
        })
    );
}

/// Two looks at the view are told apart by the facts *both* of them answered, and
/// never by one of them failing to answer: the folder behind a view is a walk out
/// through the shell's objects to a filesystem path, and a share, a slow disk or a
/// library answers nothing to that walk on some looks and a path on others. Read as
/// one key — the first fact that answered — the same place came out as `folder:–` on
/// one look and `url:–` on the next, and each was read as a change: the preview of a
/// file that never moved was taken down, the gate armed, and the same preview put
/// back a moment later, which is the blink. A look that answered nothing at all is
/// not a place to compare against — see `HoverLocation`.
#[test]
fn a_place_is_told_apart_by_the_facts_both_looks_answered() {
    let view = |folder: Option<&str>, url: Option<&str>, hwnd: isize| HoverLocation {
        folder: folder.map(str::to_string),
        search_root: None,
        location_url: url.map(str::to_string),
        view_hwnd: Some(hwnd),
    };
    let here = view(Some("D:\\Pictures"), Some("file:///D:/Pictures"), 0x1234);

    assert!(
        !hover_location_changed(&here, &view(None, Some("file:///D:/Pictures"), 0x1234)),
        "a folder the shell could not walk out this time is the same place"
    );
    assert!(
        !hover_location_changed(&view(None, Some("file:///D:/Pictures"), 0x1234), &here),
        "and so is reading it again on the next look"
    );
    assert!(
        hover_location_changed(
            &here,
            &view(Some("D:\\Videos"), Some("file:///D:/Pictures"), 0x1234)
        ),
        "another folder is another place"
    );
    assert!(
        hover_location_changed(
            &here,
            &view(Some("D:\\Pictures"), Some("file:///D:/Videos"), 0x1234)
        ),
        "and so is the same window arrived at another URL"
    );
    assert!(
        hover_location_changed(
            &here,
            &view(Some("D:\\Pictures"), Some("file:///D:/Pictures"), 0x5678)
        ),
        "and so is another view of it, in a window of its own"
    );

    let nothing = HoverLocation::default();
    assert!(
        !nothing.was_answered(),
        "a look that answered nothing is not a place to compare against"
    );
    assert!(here.was_answered(), "a look that answered one is");
}

/// The width a name is drawn at is the name's own: a longer name measures wider
/// than a short one, which is what makes the region a preview is kept off the name
/// rather than the column it sits in.
#[test]
fn a_name_is_measured_by_what_it_takes() {
    let short = drawn_name_width("aa.txt", 96).expect("a short name measures");
    let middle = drawn_name_width("mid-length-name.txt", 96).expect("a middle name measures");
    let long =
        drawn_name_width("a-very-long-file-name-here.txt", 96).expect("a long name measures");

    assert!(short > 0, "a name takes some room: {short}");
    assert!(
        short < middle && middle < long,
        "a longer name takes more room: {short} < {middle} < {long}"
    );
}

/// A name is measured in the pixels of the display it is drawn on, which is what a
/// scaled display draws it twice as wide in. Measuring one at the system's scale on
/// another display's is what covered the name: a provider that hands out no window
/// left the display unknown, and the region came out at half the name on a display
/// at 200% — see `region_display_dpi`.
#[test]
fn a_name_is_measured_at_the_scale_of_its_display() {
    let plain = drawn_name_width("mid-length-name.txt", 96).expect("a name measures");
    let scaled = drawn_name_width("mid-length-name.txt", 192).expect("a name measures");

    let ratio = scaled as f32 / plain as f32;
    assert!(
        (ratio - 2.0).abs() < 0.05,
        "twice the display draws twice the name: {plain} vs {scaled}"
    );
}

/// A display that cannot be told apart from another is no display to measure
/// against, so a name without one is left as the view reported it rather than
/// measured at a scale that is not its own — see `drawn_name_width`.
#[test]
fn a_name_without_a_display_is_not_measured() {
    assert_eq!(drawn_name_width("report.txt", 0), None);
}

/// A name with nothing in it has no width to be kept off, so the box the view
/// reported is left as it is — see `HoveredItem::name_box`.
#[test]
fn a_name_with_nothing_in_it_is_not_measured() {
    assert_eq!(drawn_name_width("", 96), None);
}

/// The name measured is the name the view shows, extension and all: what a
/// preview is kept off is the whole of what is drawn, not the stem it starts with.
#[test]
fn a_name_is_measured_with_its_extension() {
    let listed = drawn_name_width("report.txt", 96).expect("a listed name measures");
    let stem = drawn_name_width("report", 96).expect("a stem measures");

    assert!(
        listed > stem,
        "the extension takes room of its own: {stem} vs {listed}"
    );
}

/// A `Details` row's text is read as a row: the name is the `Name` column, the box
/// the item draws is the whole of its columns, and the columns drawn beside the
/// name are what says the item is a row of its view rather than a box — see
/// `text_boxes`. The boxes are the ones a row of a `Details` folder is reported
/// with, a name and the columns beside it.
#[test]
fn a_details_rows_text_is_read_as_a_row() {
    let text = text_boxes(&[
        RECT {
            left: 1160,
            top: 372,
            right: 1396,
            bottom: 391,
        },
        RECT {
            left: 1396,
            top: 372,
            right: 1540,
            bottom: 391,
        },
        RECT {
            left: 1540,
            top: 372,
            right: 1660,
            bottom: 391,
        },
        RECT {
            left: 1660,
            top: 372,
            right: 1740,
            bottom: 391,
        },
    ])
    .expect("a row that draws text");

    assert_eq!(text.name.left, 1160, "the name is the leftmost piece");
    assert_eq!(text.name.right, 1396, "which is the `Name` column");
    assert_eq!(
        text.all.right, 1740,
        "the row's text stops at its last column"
    );
    assert!(text.columns, "the columns beside the name make it a row");
}

/// A `Content` row is read the same way, though its name is drawn at 125% of the
/// icon font and its columns are not all on the name's own line: the name is still
/// the leftmost piece and the pieces drawn past it still make the item a row — see
/// `text_boxes`.
#[test]
fn a_content_rows_text_is_read_as_a_row() {
    let text = text_boxes(&[
        RECT {
            left: 1204,
            top: 349,
            right: 1504,
            bottom: 371,
        },
        RECT {
            left: 1542,
            top: 354,
            right: 1600,
            bottom: 370,
        },
        RECT {
            left: 1578,
            top: 371,
            right: 1617,
            bottom: 387,
        },
        RECT {
            left: 1836,
            top: 371,
            right: 1889,
            bottom: 387,
        },
    ])
    .expect("a row that draws text");

    assert_eq!(text.name.left, 1204, "the name is the leftmost piece");
    assert_eq!(
        text.all.right, 1889,
        "the row's text stops at its last column"
    );
    assert!(text.columns, "the details beside the name make it a row");
}

/// Text drawn *under* the name is not drawn beside it: the label under an icon is
/// one piece, and the lines a tile stacks share the room they are drawn in, so
/// neither is read as a row — the preview of such an item is placed by its box —
/// see `text_boxes`.
#[test]
fn text_under_the_name_is_not_a_column_beside_it() {
    let label = text_boxes(&[RECT {
        left: 1235,
        top: 602,
        right: 1312,
        bottom: 618,
    }])
    .expect("a label that draws text");
    assert_eq!(label.all, label.name, "the label is all of the text");
    assert!(
        !label.columns,
        "a label under an icon has nothing beside it"
    );

    let stacked = text_boxes(&[
        RECT {
            left: 1193,
            top: 346,
            right: 1384,
            bottom: 362,
        },
        RECT {
            left: 1193,
            top: 362,
            right: 1384,
            bottom: 378,
        },
        RECT {
            left: 1193,
            top: 378,
            right: 1384,
            bottom: 394,
        },
    ])
    .expect("a tile that draws text");
    assert_eq!(stacked.name.bottom, 362, "the name is the first line");
    assert!(!stacked.columns, "the lines under it are not beside it");
}

/// An item that draws no text at all has no boxes to be read, which leaves its
/// preview placed by its box — see `item_text_box`.
#[test]
fn an_item_that_draws_no_text_has_no_boxes() {
    assert!(text_boxes(&[]).is_none());
}

/// One place, spelled two ways, is one place — which is what the Shell does not
/// promise. The folder a view has open is canonicalized where that succeeds and left
/// as the view reported it where it does not, so the same folder answers the verbatim
/// form on one look and the plain one on the next. Read as it comes, that difference
/// of spelling is a difference of *place*: the preview of a file that never moved is
/// taken down on the look that answers the other spelling and put back on the one
/// after it, which is the blink the key is for. Case and a trailing separator are the
/// same difference, and a share keeps its server.
#[test]
fn one_place_spelled_two_ways_is_one_place() {
    assert_eq!(
        location_fact_key(r"\\?\D:\downloads"),
        location_fact_key(r"D:\downloads"),
        "the verbatim form names the path it is written around"
    );
    assert_eq!(
        location_fact_key(r"D:\Pictures\"),
        location_fact_key(r"D:\Pictures"),
        "a trailing separator names the place named without it"
    );
    assert_eq!(
        location_fact_key(r"D:\Pictures"),
        location_fact_key(r"d:\pictures"),
        "and Windows reads two spellings of a path as one path"
    );
    assert_eq!(
        location_fact_key(r"\\?\UNC\server\share"),
        location_fact_key(r"\\server\share"),
        "the verbatim form of a share keeps its server"
    );

    // A search's query is the query: two searches that differ in case are two
    // searches, so a fact that is one is left as it was written.
    assert_ne!(
        location_fact_key("search-ms:query=Foo"),
        location_fact_key("search-ms:query=foo")
    );
}
