use super::*;
use super::super::menu::{
    MENU_MARKER_ROOM_PIXELS, MENU_PAD_PIXELS, MENU_PANEL_MIN_PIXELS, MENU_PANEL_WIDTH_PIXELS,
};

/// The rows the card's menu carries: the two that turn a mode
/// on and off, and the one that opens the seek choices. Which
/// of them is the one in force is what the marks say.
fn three_rows() -> Vec<MenuRow> {
    vec![
        MenuRow {
            label: "Shuffle Mode".to_string(),
            mark: MenuMark::Check(true),
        },
        MenuRow {
            label: "Loop".to_string(),
            mark: MenuMark::Check(false),
        },
        MenuRow {
            label: "Seek".to_string(),
            mark: MenuMark::None,
        },
    ]
}

/// The point a menu is opened at in these tests, in the window's own
/// coordinates: somewhere that leaves room for a flyout to its right at
/// the widths the flyout cases are asked of.
fn a_point() -> (i32, i32) {
    (300, 40)
}

/// The menu opens with its top-left at the point the right-click landed,
/// and its rows are held in from its own ends by the pad the panel is
/// drawn with, so a row is never the panel's edge.
#[test]
fn the_menu_opens_at_the_point_the_click_landed() {
    let rows = three_rows();
    let popup = menu_popup_from_point((120, 40), 400, 300, 96, &rows);

    // The panel's own top-left corner is the point itself, where the
    // window leaves it room.
    assert_eq!(
        (popup.panel.left, popup.panel.top),
        (120, 40),
        "the panel's top-left is the point the click landed"
    );

    // The rows are held in from the panel's own ends by the pad the
    // panel is drawn with.
    assert!(popup.rows_top > popup.panel.top);
    let rows_bottom = popup.rows_top + rows.len() as i32 * popup.row_height;
    assert!(rows_bottom < popup.panel.bottom);

    // And the panel is inside the window it belongs to.
    assert!(popup.panel.right <= 400);
    assert!(popup.panel.bottom <= 300);
}

/// A menu asked for near an edge is held inside the window: a panel off
/// the side of a window is a row a hand cannot reach, and one off the
/// bottom is rows that cannot be read at all. The window's own scale
/// changes the panel's size, so the holding is asked of more than one.
#[test]
fn a_menu_that_would_run_off_the_window_stays_inside_it() {
    let rows = three_rows();

    // A window too narrow to hold the panel with its left at the point:
    // the panel is held at the window's own left rather than running off
    // it. The window is narrower than the panel the menu's own rows
    // measure to, which is what makes it too narrow.
    let narrow = menu_popup_from_point((90, 0), 60, 300, 96, &rows);
    assert_eq!(
        narrow.panel.left,
        0,
        "a narrow window holds the panel at its left"
    );
    assert!(narrow.panel.right <= 60);

    // A point near the window's own right edge: the panel is moved back
    // until its own right edge is the window's.
    let right = menu_popup_from_point((380, 0), 400, 300, 96, &rows);
    assert_eq!(
        right.panel.right,
        400,
        "a panel that would run off the right is moved back to the edge"
    );

    // A short window: a point near the bottom puts the panel off it, so
    // the panel is moved up until its own bottom edge is the window's.
    let short = menu_popup_from_point((10, 200), 400, 120, 96, &rows);
    assert!(
        short.panel.bottom <= 120,
        "the panel does not run off the bottom"
    );
    assert!(short.panel.top >= 0, "and it is still a panel of the window");

    // And the same two holdings at a display's own larger scale, where
    // the panel is bigger than it is at 96 DPI.
    let scaled = menu_popup_from_point((90, 200), 100, 120, 144, &rows);
    assert_eq!(
        scaled.panel.left,
        0,
        "a narrow window holds the panel at its left at 144 DPI"
    );
    assert!(scaled.panel.right <= 100);
    assert!(
        scaled.panel.bottom <= 120,
        "and a short window holds the panel at its bottom"
    );
}

/// The main menu's own panel is only as wide as its longest
/// row: a menu whose longest row is longer than the card's
/// own ("Shuffle Mode") is a wider panel at the same
/// display's own scale, and the room the longest row's label
/// asks for moves with the display's own scale the way the
/// labels' does — double the scale, and the text room is
/// roughly doubled, the way a hinted font's own rounding
/// leaves it.
#[test]
fn the_main_menu_is_only_as_wide_as_its_longest_row() {
    // A menu whose longest row is longer than the card's
    // own longest: the same three rows, the first carrying
    // a label no other row comes near.
    let longer_rows = vec![
        MenuRow {
            label: "A Longer Choice Than Any Other".to_string(),
            mark: MenuMark::Check(true),
        },
        MenuRow {
            label: "Loop".to_string(),
            mark: MenuMark::Check(false),
        },
        MenuRow {
            label: "Seek".to_string(),
            mark: MenuMark::None,
        },
    ];

    let panel_width = |rows: &[MenuRow], dpi: u32| {
        let popup = menu_popup_from_point(a_point(), 620, 300, dpi, rows);
        popup.panel.right - popup.panel.left
    };

    // The same display's own scale: the longer row's panel
    // is the wider one.
    assert!(
        panel_width(&longer_rows, 96) > panel_width(&three_rows(), 96),
        "a longer row is a wider panel at the same scale"
    );

    // Double the display's own scale: the room the longest
    // row's label asks for — the panel less the marks' room
    // and the pads — is roughly doubled.
    let text_room = |rows: &[MenuRow], dpi: u32| {
        let scale = dpi as f32 / 96.0;
        let held = text_paint::scaled(MENU_MARKER_ROOM_PIXELS as i32, scale)
            + 2 * text_paint::scaled(MENU_PAD_PIXELS as i32, scale);
        panel_width(rows, dpi) - held
    };
    let room_at_100 = text_room(&three_rows(), 96);
    let room_at_200 = text_room(&three_rows(), 192);
    assert!(
        room_at_200 > room_at_100 * 3 / 2 && room_at_200 < room_at_100 * 5 / 2,
        "double the scale roughly doubles the text room ({} to {})",
        room_at_100,
        room_at_200
    );
}

/// The main menu's panel is exactly its longest row's
/// measurement, at the display's own scale and in the theme's
/// own face, plus the room the marks stand in and the pad the
/// panel holds its rows in on either side of it — the
/// arithmetic of the width, not a size of its own.
#[test]
fn the_main_menu_panel_is_the_longest_row_plus_the_marks_and_the_pads() {
    let rows = three_rows();
    let popup = menu_popup_from_point(a_point(), 620, 300, 96, &rows);

    // The longest of the rows the panel was handed, measured
    // the way the panel's own labels are painted: the theme's
    // face at the display's own scale, on a throwaway context.
    let surface = DibSurface::create(1, 1).expect("a surface");
    let style = caption_style([0, 0, 0]);
    let scale = 96f32 / 96.0;
    let longest = rows
        .iter()
        .map(|row| measure_text(&surface, &style, &row.label, scale))
        .max()
        .expect("the menu holds rows");
    let marker_room = text_paint::scaled(MENU_MARKER_ROOM_PIXELS as i32, scale);
    let pads = 2 * text_paint::scaled(MENU_PAD_PIXELS as i32, scale);

    assert_eq!(
        popup.panel.right - popup.panel.left,
        longest + marker_room + pads,
        "the panel is the longest row plus the marks' room and the pads"
    );
}

/// A menu of the shortest rows does not collapse the panel
/// below the floor: the panel is never narrower than the
/// narrowest panel a hand can read, whatever its rows
/// measure, which is what keeps a menu of one-character
/// rows a panel rather than a sliver.
#[test]
fn a_menu_of_one_character_rows_does_not_collapse_below_the_floor() {
    let rows = vec![
        MenuRow {
            label: "I".to_string(),
            mark: MenuMark::Check(true),
        },
        MenuRow {
            label: "O".to_string(),
            mark: MenuMark::Check(false),
        },
        MenuRow {
            label: "S".to_string(),
            mark: MenuMark::None,
        },
    ];

    let popup = menu_popup_from_point(a_point(), 620, 300, 96, &rows);
    let scale = 96f32 / 96.0;
    let floor = text_paint::scaled(MENU_PANEL_MIN_PIXELS as i32, scale);
    let marker_room = text_paint::scaled(MENU_MARKER_ROOM_PIXELS as i32, scale);
    let pads = 2 * text_paint::scaled(MENU_PAD_PIXELS as i32, scale);

    assert!(
        popup.panel.right - popup.panel.left >= floor,
        "a menu of one-character rows does not collapse below the floor"
    );

    // And the floor is what holds it there: the rows alone ask
    // for less than it, so the panel is wider than its own
    // rows are.
    let surface = DibSurface::create(1, 1).expect("a surface");
    let style = caption_style([0, 0, 0]);
    let longest = rows
        .iter()
        .map(|row| measure_text(&surface, &style, &row.label, scale))
        .max()
        .expect("the menu holds rows");
    assert!(
        popup.panel.right - popup.panel.left > longest + marker_room + pads,
        "the floor held the panel open past what the rows ask for"
    );
}

/// The seek flyout's panel keeps its fixed width, whoever its
/// rows are and whatever the display's own scale: the four
/// choices and a set whose longest row is longer than any of
/// them, at two scales, are all the one fixed width — the
/// main menu's own panel is the one that is measured.
#[test]
fn the_seek_flyout_keeps_its_fixed_panel_width() {
    for (dpi, scale) in [(96u32, 1.0f32), (144, 1.5)] {
        for choices in [four_choices(), longer_flyout_rows()] {
            let menu = menu_popup_from_point(a_point(), 900, 300, dpi, &three_rows());
            let placed = menu_flyout_from_menu(&menu, 900, 300, dpi, &choices);
            assert_eq!(
                placed.popup.panel.right - placed.popup.panel.left,
                text_paint::scaled(MENU_PANEL_WIDTH_PIXELS as i32, scale),
                "the flyout's panel is the fixed width at {} DPI",
                dpi
            );
        }
    }
}

/// Both panels stay inside a window narrower than either of
/// them: the main menu's measured width is held to the
/// window's own — and the flyout's fixed one is narrowed
/// beside it — so a window too narrow for the panels is
/// still a window they are painted wholly inside, the paint
/// held to the window's bounds wherever a panel would run
/// past them.
#[test]
fn both_panels_stay_painted_inside_a_window_narrower_than_they_are() {
    let (width, height) = (100i32, 300i32);
    let palette = ChromePalette {
        background: [30, 34, 42],
        foreground: [198, 202, 210],
        accent: [86, 182, 194],
        dark: true,
    };
    let surface = DibSurface::create(width as u32, height as u32).expect("a surface");

    let menu_rows = three_rows();
    let flyout_rows = four_choices();
    let menu = menu_popup_from_point((50, 40), width, height, 96, &menu_rows);
    let placed = menu_flyout_from_menu(&menu, width, height, 96, &flyout_rows);

    // Both panels are inside the window, at the width the
    // clamps hold them to: the main menu's measured width is
    // the window's own, and the flyout's fixed one is the
    // half the window leaves beside it.
    for panel in [placed.menu.panel, placed.popup.panel] {
        assert!(
            panel.left >= 0 && panel.right <= width,
            "a panel is inside the window's own sides"
        );
        assert!(
            panel.top >= 0 && panel.bottom <= height,
            "and inside its own top and bottom"
        );
    }

    // And both panels paint, into a buffer of the window's own
    // size: what the paint lands is inside the window, which
    // is every pixel there is — a panel that would run past
    // the window's edges is clipped to them, the way the
    // placement already holds it. The page colour at a
    // pixel of the panel's own bottom pad, which no row's
    // label or mark is drawn in, is what a painted panel
    // looks like.
    let page = [palette.background[2], palette.background[1], palette.background[0]];
    let mut buffer = vec![0u8; (width * height * 4) as usize];
    // A pixel of the panel's own bottom pad, two rows above its
    // edge — the row the hairline is drawn in is the edge's own,
    // and no row's label or mark is drawn in the pad.
    let painted_inside = |buffer: &[u8], panel: RECT| {
        let middle = (panel.left + panel.right) / 2;
        let at = ((panel.bottom - 2) * width + middle) as usize * 4;
        buffer[at..at + 3] == page
    };

    paint_menu_popup(
        &mut buffer,
        width,
        &palette,
        &placed.menu,
        &menu_rows,
        None,
        &surface,
        1.0,
    );
    assert!(
        painted_inside(&buffer, placed.menu.panel),
        "the main menu is painted inside the window"
    );

    paint_menu_popup(
        &mut buffer,
        width,
        &palette,
        &placed.popup,
        &flyout_rows,
        None,
        &surface,
        1.0,
    );
    assert!(
        painted_inside(&buffer, placed.popup.panel),
        "the flyout is painted inside the window"
    );
}

/// The rows the seek flyout holds: the four seek
/// choices, in the order the menu lists them in, with
/// the one in force carrying its disc.
fn four_choices() -> Vec<MenuRow> {
    vec![
        MenuRow {
            label: "Remember".to_string(),
            mark: MenuMark::None,
        },
        MenuRow {
            label: "From the Start".to_string(),
            mark: MenuMark::None,
        },
        MenuRow {
            label: "From the Middle".to_string(),
            mark: MenuMark::None,
        },
        MenuRow {
            label: "Random".to_string(),
            mark: MenuMark::Bullet,
        },
    ]
}

/// A flyout's rows whose longest row is longer than the
/// four choices' own longest ("From the Middle"): a set
/// the flyout's fixed width is asked of, to show the
/// width does not follow its rows.
fn longer_flyout_rows() -> Vec<MenuRow> {
    vec![
        MenuRow {
            label: "Remember".to_string(),
            mark: MenuMark::None,
        },
        MenuRow {
            label: "A Choice Longer Than Any Other".to_string(),
            mark: MenuMark::None,
        },
        MenuRow {
            label: "Random".to_string(),
            mark: MenuMark::Bullet,
        },
    ]
}

/// The seek flyout hangs right of the main menu's own
/// right edge, top-aligned with the Seek row — the main
/// menu's third row — so the menu the flyout came from
/// stays wholly visible beside it, and the Seek row's
/// arrow points that way. The window is one the flyout
/// fits in beside the menu at the point, which is asked
/// of more than one width and one display's own scale,
/// because the scale is what sizes the two panels.
#[test]
fn the_seek_flyout_hangs_right_of_the_menu_on_the_seek_row() {
    let menu_rows = three_rows();
    let flyout_rows = four_choices();

    // A window the flyout fits in beside the menu at the
    // point, at the scale of a 96-DPI display and at
    // a second width: the gap the flyout is held off the
    // menu by is four logical pixels at 96 DPI.
    for width in [620i32, 700i32] {
        let placed =
            menu_flyout_from_menu(
                &menu_popup_from_point(a_point(), width, 300, 96, &menu_rows),
                width,
                300,
                96,
                &flyout_rows,
            );

        // The menu is left where the point anchored it, and
        // the flyout's left edge is a small gap right of its
        // own right edge, its top the Seek row's own top —
        // the row the flyout is hung from.
        assert_eq!(
            (placed.menu.panel.left, placed.menu.panel.top),
            a_point(),
            "the menu stays where the point anchored it"
        );
        assert_eq!(
            placed.popup.panel.left,
            placed.menu.panel.right + 4,
            "the flyout sits right of the menu's own right edge"
        );
        assert_eq!(
            placed.popup.panel.top,
            placed.menu.rows_top + 2 * placed.menu.row_height,
            "the flyout is top-aligned with the Seek row"
        );
        assert_eq!(
            placed.menu.arrow,
            MenuArrow::Right,
            "the Seek row's arrow points at the flyout on its right"
        );

        // Both panels are inside the window they belong to.
        for panel in [placed.menu.panel, placed.popup.panel] {
            assert!(panel.left >= 0, "the panel is not off the window's left");
            assert!(panel.right <= width, "nor off its right");
            assert!(panel.top >= 0, "nor off its top");
            assert!(panel.bottom <= 300, "nor off its bottom");
        }
    }

    // The same placement at a display's own larger scale,
    // where both panels are bigger than they are at 96 DPI
    // and the gap is the six logical pixels that is.
    let placed = menu_flyout_from_menu(
        &menu_popup_from_point(a_point(), 760, 300, 144, &menu_rows),
        760,
        300,
        144,
        &flyout_rows,
    );
    assert_eq!(placed.popup.panel.left, placed.menu.panel.right + 6);
    assert_eq!(
        placed.popup.panel.top,
        placed.menu.rows_top + 2 * placed.menu.row_height
    );
    assert_eq!(placed.menu.arrow, MenuArrow::Right);
    for panel in [placed.menu.panel, placed.popup.panel] {
        assert!(panel.left >= 0);
        assert!(panel.right <= 760);
        assert!(panel.top >= 0);
        assert!(panel.bottom <= 300);
    }
}

/// Where the flyout does not fit right of the menu, it flips to
/// the menu's own left, a small gap off its left edge with the same top
/// alignment — the standard Windows behavior — and the Seek row's
/// arrow flips with it. The menu is not moved to make room on the
/// right; it stays where the point anchored it, at whatever width
/// its own rows measure to.
#[test]
fn the_seek_flyout_flips_left_where_it_does_not_fit_right() {
    let menu_rows = three_rows();
    let flyout_rows = four_choices();

    // A window the flyout does not fit in beside the menu
    // where the point anchored it: the menu's own right edge
    // and the gap and the flyout's fixed width together run
    // past the window's own right, so the flyout goes to the
    // menu's left.
    for (width, dpi, gap, panel_width) in [(400i32, 96u32, 4i32, 150i32), (500, 144, 6, 225)] {
        let placed = menu_flyout_from_menu(
            &menu_popup_from_point(a_point(), width, 300, dpi, &menu_rows),
            width,
            300,
            dpi,
            &flyout_rows,
        );

        assert_eq!(
            placed.menu.arrow,
            MenuArrow::Left,
            "the Seek row's arrow points at the flyout on its left"
        );
        assert_eq!(
            placed.popup.panel.left,
            placed.menu.panel.left - gap - panel_width,
            "the flyout sits a gap left of the menu's own left edge"
        );
        assert!(
            placed.popup.panel.right <= placed.menu.panel.left,
            "and it does not cover the menu"
        );
        assert_eq!(
            placed.popup.panel.top,
            placed.menu.rows_top + 2 * placed.menu.row_height,
            "the flipped flyout is top-aligned with the Seek row"
        );

        // The menu is left where the point anchored it — the
        // flyout's placement does not move it to make room on
        // the right, whatever width the menu's own rows
        // measure to — and both panels are inside the window.
        assert_eq!(
            placed.menu.panel,
            menu_popup_from_point(a_point(), width, 300, dpi, &menu_rows).panel,
            "the menu did not move"
        );
        for panel in [placed.menu.panel, placed.popup.panel] {
            assert!(panel.left >= 0, "a panel is off the window's left");
            assert!(panel.right <= width, "a panel is off the window's right");
            assert!(panel.top >= 0 && panel.bottom <= 300);
        }
    }
}

/// The Seek row carries its arrow whenever the menu is up, whether
/// its flyout is up or not: the arrow is a property of the main
/// menu's own panel, answered from the same tier the flyout's own
/// placement asks (see `menu_flyout_tier`), so the row says it opens
/// a submenu before the submenu is there — the way every submenu
/// door reads. Asked of every tier a window gives the pair, at more
/// than one display's own scale, which is what sizes the panels.
#[test]
fn the_seek_row_carries_its_arrow_while_the_flyout_is_down() {
    let rows = three_rows();

    // The tiers, at the widths that give them: a window the flyout
    // fits in right of the menu at the point, a window wide enough
    // only for the flyout to the menu's left, and a window too narrow
    // for the two panels at their own width, which narrows them — the
    // narrowed pair pointing the arrow right, the flyout being at the
    // window's right half.
    let placements = [
        (620i32, 300i32, 96u32, MenuArrow::Right), // the flyout fits right
        (400, 300, 96, MenuArrow::Left), // the flyout flips left
        (250, 300, 96, MenuArrow::Right), // the two are narrowed
        (760, 300, 144, MenuArrow::Right), // fits right, larger scale
        (500, 300, 144, MenuArrow::Left), // flips left, larger scale
        (250, 300, 144, MenuArrow::Right), // narrowed, larger scale
    ];

    for (width, height, dpi, arrow) in placements {
        let popup = menu_popup_from_point(a_point(), width, height, dpi, &rows);
        assert_eq!(
            popup.arrow, arrow,
            "the Seek row's arrow points the way its flyout opens, with the flyout down ({}x{} at {} DPI)",
            width, height, dpi
        );

        // The panel itself is the main menu's own placement: the anchor
        // held inside the window, and the whole panel inside it — the
        // tier's narrowing of the main menu is a flyout-up concern, not
        // this one's.
        assert!(popup.panel.left >= 0 && popup.panel.right <= width);
        assert!(popup.panel.top >= 0 && popup.panel.bottom <= height);
    }
}

/// The arrow the main menu carries while the flyout is down is the side
/// the flyout opens on when it does open: the arrow and the flyout's
/// placement are one computation, the same tier asked of the same panel
/// (`menu_flyout_tier`), so the arrow cannot disagree with the flyout —
/// and it does not move when the flyout comes up. Asked of every tier a
/// window gives the pair.
#[test]
fn the_arrow_while_the_flyout_is_down_is_the_side_the_flyout_opens_on() {
    let menu_rows = three_rows();
    let flyout_rows = four_choices();

    let placements = [
        (620i32, 300i32, 96u32), // the flyout fits right
        (400, 300, 96), // the flyout flips left
        (250, 300, 96), // the two are narrowed
        (760, 300, 144), // fits right, larger scale
        (500, 300, 144), // flips left, larger scale
        (250, 300, 144), // narrowed, larger scale
    ];

    for (width, height, dpi) in placements {
        // The panel as the main menu's own placement answers it, with the
        // flyout down, and the arrow it carries.
        let down = menu_popup_from_point(a_point(), width, height, dpi, &menu_rows);
        assert_ne!(
            down.arrow,
            MenuArrow::None,
            "the Seek row carries an arrow while the menu is up ({}x{} at {} DPI)",
            width, height, dpi
        );

        // The same panel with the flyout up: the two placed together,
        // the flyout beside the menu.
        let placed = menu_flyout_from_menu(&down, width, height, dpi, &flyout_rows);

        assert_eq!(
            placed.menu.arrow, down.arrow,
            "the arrow does not move when the flyout comes up ({}x{} at {} DPI)",
            width, height, dpi
        );

        // And the flyout's panel is on the side the arrow points: right
        // of the menu where the arrow points right, left of it where the
        // arrow points left.
        match down.arrow {
            MenuArrow::Right => assert!(
                placed.popup.panel.left >= placed.menu.panel.right,
                "the flyout opens right of the menu, the arrow's side ({}x{} at {} DPI)",
                width, height, dpi
            ),
            MenuArrow::Left => assert!(
                placed.popup.panel.right <= placed.menu.panel.left,
                "the flyout opens left of the menu, the arrow's side ({}x{} at {} DPI)",
                width, height, dpi
            ),
            MenuArrow::None => {
                unreachable!("the Seek row always carries an arrow while the menu is up")
            }
        }
    }
}

/// Both panels stay inside the window at every width: where the
/// flyout fits on neither side of the menu it is anchored at — a
/// window too narrow for the two panels at their own width — the two
/// are narrowed to half the room the window leaves between its own
/// edges, the menu at the window's left edge and the flyout right of
/// it, the arrow pointing at it. A window too short to hold the
/// flyout below the Seek row is a flyout held at the window's own
/// bottom, the same holding the menu's own panel has. Each holding is
/// asked of more than one display's own scale.
#[test]
fn both_panels_stay_inside_the_window_at_every_width() {
    let menu_rows = three_rows();
    let flyout_rows = four_choices();

    // A window too narrow for the flyout on either side of the
    // menu at the point: the two are narrowed to half the room
    // the window leaves between its own edges, the menu at the
    // window's left edge and the flyout right of it.
    let split = menu_flyout_from_menu(
        &menu_popup_from_point(a_point(), 250, 300, 96, &menu_rows),
        250,
        300,
        96,
        &flyout_rows,
    );
    assert_eq!(
        split.menu.panel.left,
        0,
        "a window too narrow for the two holds the menu at its left"
    );
    assert_eq!(split.popup.panel.right, 250, "and the flyout at its right");
    assert_eq!(
        split.popup.panel.left,
        split.menu.panel.right + 4,
        "the narrowed flyout is still right of the narrowed menu"
    );
    assert!(split.menu.panel.right < split.popup.panel.left);
    assert_eq!(split.menu.arrow, MenuArrow::Right);

    // And the same narrowing at a display's own larger scale.
    let split = menu_flyout_from_menu(
        &menu_popup_from_point(a_point(), 250, 300, 144, &menu_rows),
        250,
        300,
        144,
        &flyout_rows,
    );
    assert_eq!(split.menu.panel.left, 0);
    assert_eq!(split.popup.panel.right, 250);
    assert_eq!(split.popup.panel.left, split.menu.panel.right + 6);
    assert!(split.menu.panel.right < split.popup.panel.left);

    // A window too short to hold the flyout below the Seek
    // row: the flyout is held at the window's own bottom,
    // moved up from the row it is aligned with, and the menu
    // is still inside the window.
    let short = menu_flyout_from_menu(
        &menu_popup_from_point(a_point(), 600, 120, 96, &menu_rows),
        600,
        120,
        96,
        &flyout_rows,
    );
    assert_eq!(
        short.popup.panel.bottom,
        120,
        "a short window holds the flyout at its bottom"
    );
    assert!(
        short.popup.panel.top < short.menu.rows_top + 2 * short.menu.row_height,
        "the flyout is moved up from the Seek row to stay inside"
    );
    assert!(short.popup.panel.top >= 0);
    assert!(
        short.menu.panel.bottom <= 120,
        "and the menu is still inside the short window"
    );
}

/// While the flyout is up, the menu it came from stays
/// wholly visible and uncovered: every row of the menu is
/// inside the window, and the flyout's panel does not reach
/// into the menu's own — the flyout is beside the menu, not
/// over it. Asked of the placements a window gives the pair:
/// the flyout at the menu's side, the menu pulled left to
/// make the flyout room, and the two narrowed to fit a
/// window too small for them at their own width.
#[test]
fn every_menu_row_stays_visible_and_uncovered_while_the_flyout_is_up() {
    let menu_rows = three_rows();
    let flyout_rows = four_choices();

    let placements = [
        (620i32, 300i32, 96u32), // the flyout to the menu's right
        (400, 300, 96),          // flipped to the menu's left
        (250, 300, 96),          // the two narrowed
        (500, 300, 144),         // flipped left at a larger scale
        (250, 300, 144),         // narrowed at a larger scale
    ];

    for (width, height, dpi) in placements {
        let placed = menu_flyout_from_menu(
            &menu_popup_from_point(a_point(), width, height, dpi, &menu_rows),
            width,
            height,
            dpi,
            &flyout_rows,
        );
        let menu = placed.menu;
        let flyout = placed.popup;

        // The flyout is beside the menu, not over it: its panel
        // lies wholly off one of the menu's own sides, so no row
        // of the menu is covered — the flyout is to the menu's
        // right where it fits there and to its left where it does
        // not (see `menu_flyout_from_menu`).
        assert!(
            flyout.panel.left >= menu.panel.right || flyout.panel.right <= menu.panel.left,
            "the flyout does not cover the menu ({}x{} at {} DPI)",
            width,
            height,
            dpi
        );

        // Every row of the menu is inside the window: the
        // rows' own band, from the first row's top to the last
        // row's bottom, held between the window's own edges
        // with the panel it is drawn in.
        let rows_bottom = menu.rows_top + menu.rows * menu.row_height;
        assert!(
            menu.rows_top >= 0 && rows_bottom <= height,
            "the menu's rows are inside the window ({}x{} at {} DPI)",
            width,
            height,
            dpi
        );
        assert!(menu.panel.left >= 0 && menu.panel.right <= width);
        for index in 0..menu.rows {
            let row_top = menu.rows_top + index * menu.row_height;
            assert!(
                row_top >= 0 && row_top + menu.row_height <= height,
                "menu row {index} is inside the window ({}x{} at {} DPI)",
                width,
                height,
                dpi
            );
        }

        // And the flyout is inside the window too, which is
        // what keeps it reachable.
        assert!(flyout.panel.left >= 0 && flyout.panel.right <= width);
        assert!(flyout.panel.top >= 0 && flyout.panel.bottom <= height);
    }
}

/// A point answers a row only inside the rows: the pad the panel is
/// drawn with above and below them is no row at all, and neither is
/// anywhere outside the panel, above it or below it — a press there is
/// a press outside the menu, which is what closes it.
#[test]
fn a_point_is_a_row_only_inside_the_rows() {
    let popup = menu_popup_from_point(a_point(), 400, 300, 96, &three_rows());

    let middle_x = (popup.panel.left + popup.panel.right) / 2;
    let row_middle = |index: i32| {
        popup.rows_top + index * popup.row_height + popup.row_height / 2
    };

    assert_eq!(menu_row_at(&popup, middle_x, row_middle(0)), Some(0));
    assert_eq!(menu_row_at(&popup, middle_x, row_middle(1)), Some(1));
    assert_eq!(menu_row_at(&popup, middle_x, row_middle(2)), Some(2));
    // The last row's own bottom pixel is still that row.
    assert_eq!(
        menu_row_at(&popup, middle_x, popup.rows_top + 3 * popup.row_height - 1),
        Some(2)
    );

    // The pad above the rows and the pad below them.
    assert_eq!(menu_row_at(&popup, middle_x, popup.rows_top - 1), None);
    assert_eq!(
        menu_row_at(&popup, middle_x, popup.rows_top + 3 * popup.row_height),
        None
    );
    // Outside the panel's own edges, and above and below the panel.
    assert_eq!(
        menu_row_at(&popup, popup.panel.left - 1, row_middle(0)),
        None
    );
    assert_eq!(menu_row_at(&popup, popup.panel.right, row_middle(0)), None);
    assert_eq!(menu_row_at(&popup, middle_x, popup.panel.top - 1), None);
    assert_eq!(menu_row_at(&popup, middle_x, popup.panel.bottom), None);
}

/// A row's mark is painted in the room the labels are held away
/// from, and each kind of mark is its own art there: a check row the
/// check alone — no box — where the setting is on and nothing at all
/// where it is off, and a choice of a seek page its disc. The row that
/// carries no mark carries nothing there at all.
#[test]
fn a_check_row_carries_a_check_alone_and_a_choice_its_disc() {
    let (width, height) = (400i32, 300i32);

    // What is behind the menu: a flat colour the theme holds
    // nothing of, so the panel and what is drawn in it are
    // told apart from it.
    let backdrop = [90u8, 60, 30];
    let mut buffer = vec![0u8; (width * height * 4) as usize];
    for pixel in buffer.as_chunks_mut::<4>().0 {
        pixel[0] = backdrop[0];
        pixel[1] = backdrop[1];
        pixel[2] = backdrop[2];
        pixel[3] = 255;
    }

    let palette = ChromePalette {
        background: [30, 34, 42],
        foreground: [198, 202, 210],
        accent: [86, 182, 194],
        dark: true,
    };
    let surface = DibSurface::create(width as u32, height as u32).expect("a surface");
    let rows = vec![
        MenuRow {
            label: "Shuffle Mode".to_string(),
            mark: MenuMark::Check(true),
        },
        MenuRow {
            label: "Loop".to_string(),
            mark: MenuMark::Check(false),
        },
        MenuRow {
            label: "Seek".to_string(),
            mark: MenuMark::None,
        },
        MenuRow {
            label: "From the Start".to_string(),
            mark: MenuMark::Bullet,
        },
    ];
    let popup = menu_popup_from_point(a_point(), width, height, 96, &rows);
    paint_menu_popup(&mut buffer, width, &palette, &popup, &rows, None, &surface, 1.0);

    // The room the marks stand in: the pad the panel holds its
    // rows in and the column the marks are centred in, inside
    // the hairline that keeps the panel off what is behind it
    // and short of the labels it holds.
    let marks = popup.panel.left + 1..popup.panel.left + 22;

    /// The pixels of a band of the buffer that are neither the
    /// panel's own page nor what was behind it, which is what
    /// a mark being drawn means.
    fn ink_of(
        buffer: &[u8],
        width: i32,
        rows: std::ops::Range<i32>,
        columns: std::ops::Range<i32>,
        page: [u8; 3],
        behind: [u8; 3],
    ) -> Vec<(i32, i32)> {
        let mut ink = Vec::new();
        for y in rows {
            for x in columns.clone() {
                let at = ((y * width + x) as usize) * 4;
                if buffer[at..at + 3] != page && buffer[at..at + 3] != behind {
                    ink.push((x, y));
                }
            }
        }
        ink
    }

    let page = [palette.background[2], palette.background[1], palette.background[0]];
    let behind = [backdrop[2], backdrop[1], backdrop[0]];
    let row_ink = |index: usize| {
        let top = popup.rows_top + index as i32 * popup.row_height;
        ink_of(
            &buffer,
            width,
            top..top + popup.row_height,
            marks.clone(),
            page,
            behind,
        )
    };

    // The mark of a row is centred in the room the labels are
    // held away from, which is where its art is measured from.
    let centre = |index: usize| {
        (
            popup.panel.left as f32 + 6.0 + 8.0,
            popup.rows_top as f32
                + index as f32 * popup.row_height as f32
                + popup.row_height as f32 / 2.0
                + 0.5,
        )
    };
    let near_to = |ink: &[(i32, i32)], at: (f32, f32), within: f32| {
        ink.iter().all(|&(x, y)| {
            let dx = x as f32 + 0.5 - at.0;
            let dy = y as f32 + 0.5 - at.1;
            (dx * dx + dy * dy).sqrt() <= within
        })
    };

    // The checked row: a check mark in the room and nothing
    // around it — the box is gone. Its ink is held near the
    // mark's own centre, which is what a check alone is rather
    // than the square that used to be drawn around one.
    let checked = row_ink(0);
    assert!(!checked.is_empty(), "a checked row carries its check");
    assert!(
        near_to(&checked, centre(0), 5.0),
        "the check alone stands in the room, with no box around it"
    );
    // And the check is a check and not a dot: its ink lies on
    // both sides of the mark's own centre column.
    assert!(
        checked.iter().any(|&(x, _)| (x as f32 + 0.5) < centre(0).0)
            && checked.iter().any(|&(x, _)| (x as f32 + 0.5) > centre(0).0),
        "the check reaches across its own middle"
    );

    // The row whose setting is off: nothing at all in the room —
    // no empty box, no mark.
    assert!(
        row_ink(1).is_empty(),
        "a setting that is off carries nothing at all"
    );

    // The row that carries no mark: nothing at all in the room.
    assert!(row_ink(2).is_empty(), "the Seek row carries no mark");

    // The disc of a choice in force: ink held near the mark's
    // own centre and nowhere else, which is what a disc is
    // rather than a box.
    let bullet = row_ink(3);
    assert!(!bullet.is_empty(), "a choice in force carries its disc");
    assert!(
        near_to(&bullet, centre(3), 4.0),
        "the mark of a choice is a disc, held near the centre"
    );
}

/// A row under the pointer is washed: the band the pointer is on is
/// painted the theme's own lift, and no other row is — which is what
/// tells a hand where on the menu it is about to land.
#[test]
fn the_row_under_the_pointer_is_washed() {
    let (width, height) = (400i32, 300i32);
    let backdrop = [90u8, 60, 30];
    let palette = ChromePalette {
        background: [30, 34, 42],
        foreground: [198, 202, 210],
        accent: [86, 182, 194],
        dark: true,
    };
    let surface = DibSurface::create(width as u32, height as u32).expect("a surface");

    // The row under the pointer is the second one, and the menu is
    // painted with it asked for.
    let hovered = 1usize;
    let rows = three_rows();
    let popup = menu_popup_from_point(a_point(), width, height, 96, &rows);

    let band = |index: usize| {
        let top = popup.rows_top + index as i32 * popup.row_height;
        (top, top + popup.row_height)
    };
    let page = [palette.background[2], palette.background[1], palette.background[0]];

    // A pixel of a row that is the row's own colour and nothing
    // else: the far pad at the row's own left, left of the marker
    // column, where no label or mark is drawn.
    let background_of = |buffer: &[u8], index: usize| {
        let (top, _) = band(index);
        let at = (((top + 2) * width) as usize + (popup.panel.left + 2) as usize) * 4;
        [buffer[at], buffer[at + 1], buffer[at + 2]]
    };

    let mut buffer = vec![0u8; (width * height * 4) as usize];
    for pixel in buffer.as_chunks_mut::<4>().0 {
        pixel[0] = backdrop[0];
        pixel[1] = backdrop[1];
        pixel[2] = backdrop[2];
        pixel[3] = 255;
    }
    paint_menu_popup(
        &mut buffer,
        width,
        &palette,
        &popup,
        &rows,
        Some(hovered),
        &surface,
        1.0,
    );

    assert_ne!(
        background_of(&buffer, hovered),
        page,
        "the row under the pointer is washed, not the panel's own page"
    );
    for other in [0usize, 2] {
        assert_eq!(
            background_of(&buffer, other),
            page,
            "a row the pointer is not on is the panel's own page"
        );
    }
}

/// The Seek row carries an arrow at its own right, pointing toward the
/// side the flyout sits on: a filled triangle whose ink is on that row
/// and nowhere else, and whose point is the side the placement chose.
#[test]
fn the_seek_row_carries_an_arrow_pointing_at_the_flyout() {
    let (width, height) = (620i32, 300i32);
    let backdrop = [90u8, 60, 30];
    let palette = ChromePalette {
        background: [30, 34, 42],
        foreground: [198, 202, 210],
        accent: [86, 182, 194],
        dark: true,
    };
    let surface = DibSurface::create(width as u32, height as u32).expect("a surface");

    // The same menu painted twice: once with no arrow at all, once
    // with the arrow pointing right and once pointing left. The ink
    // in the band at the row's own right is the arrow's.
    let paint = |arrow: MenuArrow| -> Vec<u8> {
        let mut popup = menu_popup_from_point(a_point(), width, height, 96, &three_rows());
        popup.arrow = arrow;
        let mut buffer = vec![0u8; (width * height * 4) as usize];
        for pixel in buffer.as_chunks_mut::<4>().0 {
            pixel[0] = backdrop[0];
            pixel[1] = backdrop[1];
            pixel[2] = backdrop[2];
            pixel[3] = 255;
        }
        paint_menu_popup(
            &mut buffer,
            width,
            &palette,
            &popup,
            &three_rows(),
            None,
            &surface,
            1.0,
        );
        buffer
    };

    let base = menu_popup_from_point(a_point(), width, height, 96, &three_rows());
    let row_top = base.rows_top + 2 * base.row_height;
    // The band at the row's own right that the arrow stands in: from a
    // little inside the panel's right edge to a little inside it again,
    // short of the hairline the panel is outlined with.
    let columns = base.panel.right - 25..base.panel.right - 3;
    let page = [palette.background[2], palette.background[1], palette.background[0]];
    let behind = [backdrop[2], backdrop[1], backdrop[0]];

    // The ink of the arrow in the row that carries it: the pixels of
    // the seek row's right band that are neither the page nor the
    // backdrop.
    let arrow_ink = |buffer: &[u8]| -> Vec<(i32, i32)> {
        let mut ink = Vec::new();
        for y in row_top..row_top + base.row_height {
            for x in columns.clone() {
                let at = ((y * width + x) as usize) * 4;
                if buffer[at..at + 3] != page && buffer[at..at + 3] != behind {
                    ink.push((x, y));
                }
            }
        }
        ink
    };

    let plain = paint(MenuArrow::None);
    let right = paint(MenuArrow::Right);
    let left = paint(MenuArrow::Left);

    assert!(
        arrow_ink(&plain).is_empty(),
        "a row that opens nothing carries no arrow"
    );
    assert!(
        !arrow_ink(&right).is_empty(),
        "the Seek row carries an arrow where the flyout is to its right"
    );
    assert!(
        !arrow_ink(&left).is_empty(),
        "and an arrow where the flyout is to its left"
    );

    // Which way it points: the two are the same triangle turned about
    // the mark's own middle, so the ink of the left-pointing one is
    // further right than the right-pointing one's — its point is the
    // left edge and its base the right for one, and the other way
    // round for the other.
    let mean = |ink: &[(i32, i32)]| {
        ink.iter().map(|&(x, _)| x as f64).sum::<f64>() / ink.len() as f64
    };
    let right_ink = arrow_ink(&right);
    let left_ink = arrow_ink(&left);
    assert!(
        mean(&left_ink) > mean(&right_ink),
        "the left-pointing arrow's ink sits further right than the right-pointing one's"
    );
}

/// The flyout's panel is absent while the flyout is down: the
/// menu paints its own panel alone, so the room the flyout would
/// open into — the gap right of the main menu's own right edge
/// and beyond it, over the Seek row's band — holds nothing of the
/// menu, which is what a panel that is not carried is rather than
/// one painted empty. The arrow the row carries sits inside the
/// panel, not out in that room.
#[test]
fn the_flyout_s_panel_is_not_painted_while_the_flyout_is_down() {
    let (width, height) = (620i32, 300i32);

    // What is behind the menu: a flat colour the theme holds
    // nothing of, so anything of the menu is told apart from it.
    let backdrop = [90u8, 60, 30];
    let palette = ChromePalette {
        background: [30, 34, 42],
        foreground: [198, 202, 210],
        accent: [86, 182, 194],
        dark: true,
    };
    let surface = DibSurface::create(width as u32, height as u32).expect("a surface");

    let rows = three_rows();
    let popup = menu_popup_from_point(a_point(), width, height, 96, &rows);
    let mut buffer = vec![0u8; (width * height * 4) as usize];
    for pixel in buffer.as_chunks_mut::<4>().0 {
        pixel[0] = backdrop[0];
        pixel[1] = backdrop[1];
        pixel[2] = backdrop[2];
        pixel[3] = 255;
    }
    paint_menu_popup(&mut buffer, width, &palette, &popup, &rows, None, &surface, 1.0);

    // The room the flyout would open into while it is down: from
    // the gap the flyout is held off the menu by — a small gap
    // right of the main menu's own right edge — to the window's
    // own right edge, over the Seek row's band, the band the
    // flyout is top-aligned with when it is up. A window the
    // flyout fits in, so the room is empty because nothing is
    // carried into it, not because there is no room for it.
    let seek_top = popup.rows_top + 2 * popup.row_height;
    for y in seek_top..seek_top + popup.row_height {
        for x in popup.panel.right + 4..width {
            let at = ((y * width + x) as usize) * 4;
            assert_eq!(
                &buffer[at..at + 3],
                &backdrop,
                "the flyout's panel is not painted while the flyout is down"
            );
        }
    }
}

/// A menu is painted as the theme's own panel over what is behind it:
/// the page colour, a hairline around it, a shadow under it, and the
/// rows' labels written in the ink — the same way a tool tip is drawn,
/// because a menu is the same kind of thing floating over the media.
#[test]
fn a_menu_is_painted_as_a_panel_of_the_theme_over_what_is_behind_it() {
    let (width, height) = (400i32, 300i32);

    // What is behind the menu: a flat colour the theme holds nothing
    // of, so the panel and its text are told apart from it.
    let backdrop = [90u8, 60, 30];
    let mut buffer = vec![0u8; (width * height * 4) as usize];
    for pixel in buffer.as_chunks_mut::<4>().0 {
        pixel[0] = backdrop[0];
        pixel[1] = backdrop[1];
        pixel[2] = backdrop[2];
        pixel[3] = 255;
    }

    let palette = ChromePalette {
        background: [30, 34, 42],
        foreground: [198, 202, 210],
        accent: [86, 182, 194],
        dark: true,
    };
    let surface = DibSurface::create(width as u32, height as u32).expect("a surface");
    let rows = three_rows();
    let popup = menu_popup_from_point(a_point(), width, height, 96, &rows);
    paint_menu_popup(&mut buffer, width, &palette, &popup, &rows, None, &surface, 1.0);

    // The panel is the theme's own page colour, drawn over what is
    // behind it: the bottom pad is the page and opaque, not the
    // backdrop it covers. (The pad rather than the panel's middle,
    // because the middle is a row's own room, and a label's ink is
    // allowed to be there.)
    let pad = (popup.panel.bottom - 2, (popup.panel.left + popup.panel.right) / 2);
    let at = ((pad.0 * width + pad.1) as usize) * 4;
    assert_eq!(
        &buffer[at..at + 3],
        &[palette.background[2], palette.background[1], palette.background[0]],
        "the panel is the theme's page colour"
    );
    assert_eq!(buffer[at + 3], 255, "and it is opaque there");

    // Each row's label is written in the ink: somewhere in the row's
    // own band is a pixel that is neither the page nor what was behind
    // the panel, which is what a label being drawn means.
    let page = [palette.background[2], palette.background[1], palette.background[0]];
    let behind = [backdrop[2], backdrop[1], backdrop[0]];
    for index in 0..rows.len() {
        let row_top = popup.rows_top + index as i32 * popup.row_height;
        let mut ink = 0;
        for y in row_top..row_top + popup.row_height {
            for x in popup.panel.left..popup.panel.right {
                let at = ((y * width + x) as usize) * 4;
                if buffer[at..at + 3] != page && buffer[at..at + 3] != behind {
                    ink += 1;
                }
            }
        }
        assert!(ink > 0, "row {index} is written in the ink");
    }
}
