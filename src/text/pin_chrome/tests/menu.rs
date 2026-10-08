use super::*;

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

/// The box the card's own arithmetic answers for the gear, in
/// the window's own coordinates: the box a press is answered
/// against, which reaches the window's own top edge above the
/// drawn glyph and down to the bottom of the band the window
/// buttons stand in.
fn the_gear() -> RECT {
    RECT {
        left: 346,
        top: 0,
        right: 362,
        bottom: 21,
    }
}

/// The menu opens downward from the gear's own button, with
/// its right edge tucked to the gear's own left edge a small
/// gap off it — the placement that leaves a flyout room to
/// the menu's right — and the panel hangs below the band the
/// gear stands in rather than over it.
#[test]
fn the_menu_opens_down_from_the_button_with_its_right_edge_tucked_to_its_left() {
    let gear = the_gear();
    let rows = three_rows();
    let popup = menu_popup_from_button(gear, 400, 300, 96, &rows);

    // One small gap off the gear's own left edge, at the
    // scale of a 96-DPI display, so the panel's right edge
    // is tucked to the gear rather than hung off its right.
    assert_eq!(popup.panel.right, gear.left - 4);
    assert!(
        popup.panel.top >= gear.bottom,
        "the panel hangs below the band the gear stands in, not over it"
    );

    // The rows are held in from the panel's own ends by the pad the
    // panel is drawn with, so a row is never the panel's edge.
    assert!(popup.rows_top > popup.panel.top);
    let rows_bottom = popup.rows_top + rows.len() as i32 * popup.row_height;
    assert!(rows_bottom < popup.panel.bottom);

    // And the panel is inside the window it belongs to.
    assert!(popup.panel.right <= 400);
    assert!(popup.panel.bottom <= 300);
}

/// A menu that would run off the window is held inside it: a
/// panel off the side of a window is a row a hand cannot
/// reach, and one off the bottom is rows that cannot be read
/// at all. The window's own scale changes the panel's size,
/// so the holding is asked of more than one.
#[test]
fn a_menu_that_would_run_off_the_window_stays_inside_it() {
    let rows = three_rows();

    // A narrow window: the panel is wider than the room the
    // gear's own left edge leaves to the left, so the panel
    // is narrowed to the window and held at its left edge
    // rather than running off it.
    let narrow = menu_popup_from_button(
        RECT {
            left: 46,
            top: 0,
            right: 62,
            bottom: 21,
        },
        100,
        300,
        96,
        &rows,
    );
    assert_eq!(
        narrow.panel.left,
        0,
        "a narrow window holds the panel at its left"
    );
    assert!(narrow.panel.right <= 100);

    // A short window: the gear sits near the bottom and the
    // panel cannot hang below it inside the window, so the
    // panel is held at the window's own bottom instead.
    let short = menu_popup_from_button(
        RECT {
            left: 346,
            top: 0,
            right: 362,
            bottom: 113,
        },
        400,
        120,
        96,
        &rows,
    );
    assert!(
        short.panel.bottom <= 120,
        "the panel does not run off the bottom"
    );
    assert!(short.panel.top >= 0, "and it is still a panel of the window");

    // And the same two holdings at a display's own larger
    // scale, where the panel is bigger than it is at 96 DPI:
    // the room the gear's left edge leaves is smaller than
    // the panel, and the window is shorter than the panel
    // below it.
    let scaled = menu_popup_from_button(
        RECT {
            left: 46,
            top: 0,
            right: 62,
            bottom: 21,
        },
        100,
        120,
        144,
        &rows,
    );
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

/// The seek flyout hangs right of the main menu's own
/// right edge, top-aligned with the Seek row — the main
/// menu's third row — so the menu the flyout came from
/// stays wholly visible beside it. The window is one the
/// flyout fits in beside the menu at the gear's side,
/// which is asked of more than one width and one display's
/// own scale, because the scale is what sizes the two
/// panels.
#[test]
fn the_seek_flyout_hangs_right_of_the_menu_on_the_seek_row() {
    let menu_rows = three_rows();
    let flyout_rows = four_choices();

    // A window the flyout fits in beside the menu at the
    // gear's side, at the scale of a 96-DPI display and at
    // a second width, where the gap the flyout is held off
    // the menu by is the same four logical pixels the menu
    // is held off the gear by.
    for width in [520i32, 600i32] {
        let placed =
            menu_flyout_from_menu(
                &menu_popup_from_button(the_gear(), width, 300, 96, &menu_rows),
                width,
                300,
                96,
                &flyout_rows,
            );

        // The flyout's left edge is a small gap right of the
        // main menu's own right edge, and its top is the Seek
        // row's own top — the row the flyout is hung from.
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
        &menu_popup_from_button(the_gear(), 700, 300, 144, &menu_rows),
        700,
        300,
        144,
        &flyout_rows,
    );
    assert_eq!(placed.popup.panel.left, placed.menu.panel.right + 6);
    assert_eq!(
        placed.popup.panel.top,
        placed.menu.rows_top + 2 * placed.menu.row_height
    );
    for panel in [placed.menu.panel, placed.popup.panel] {
        assert!(panel.left >= 0);
        assert!(panel.right <= 700);
        assert!(panel.top >= 0);
        assert!(panel.bottom <= 300);
    }
}

/// Both panels stay inside the window at every width: the
/// button the menu hangs from stands in the window's top
/// corner, where a menu tucked to it leaves a flyout no
/// room but the corner's own — so a window too narrow for
/// the flyout beside the menu there is a menu pulled left,
/// off the button, until the flyout fits beside it at the
/// window's own right edge, and a window too narrow for
/// the two panels at their own width is the two of them
/// narrowed to half the room the window leaves between its
/// own edges. A window too short to hold the flyout below
/// the Seek row is a flyout held at the window's own
/// bottom, the same holding the menu's own panel has. Each
/// holding is asked of more than one display's own scale.
#[test]
fn both_panels_stay_inside_the_window_at_every_width() {
    let menu_rows = three_rows();
    let flyout_rows = four_choices();

    // A window the flyout runs off the right edge of at the
    // menu's tucked place, but one two panels fit in side by
    // side: the menu is pulled left until the flyout fits
    // beside it at the window's right edge, and the menu
    // stays inside the window and uncovered.
    let pulled = menu_flyout_from_menu(
        &menu_popup_from_button(the_gear(), 400, 300, 96, &menu_rows),
        400,
        300,
        96,
        &flyout_rows,
    );
    assert_eq!(
        pulled.popup.panel.right,
        400,
        "the flyout is held at the window's right edge"
    );
    assert_eq!(
        pulled.popup.panel.left,
        pulled.menu.panel.right + 4,
        "and it is still right of the menu"
    );
    assert!(
        pulled.menu.panel.left >= 0,
        "the menu pulled left is still inside the window"
    );
    assert!(
        pulled.menu.panel.right < pulled.popup.panel.left,
        "and the pulled menu is not covered by the flyout"
    );

    // The same holding at a display's own larger scale, where
    // the panels are bigger than the window can hold side by
    // side at the button's side.
    let pulled = menu_flyout_from_menu(
        &menu_popup_from_button(the_gear(), 500, 300, 144, &menu_rows),
        500,
        300,
        144,
        &flyout_rows,
    );
    assert_eq!(pulled.popup.panel.right, 500);
    assert_eq!(pulled.popup.panel.left, pulled.menu.panel.right + 6);
    assert!(pulled.menu.panel.left >= 0);
    assert!(pulled.menu.panel.right < pulled.popup.panel.left);

    // A window too narrow for the two panels at their own
    // width: the two are narrowed to half the room the
    // window leaves between its own edges, the menu at the
    // window's left edge and the flyout right of it.
    let split = menu_flyout_from_menu(
        &menu_popup_from_button(the_gear(), 250, 300, 96, &menu_rows),
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
    assert_eq!(
        split.popup.panel.right,
        250,
        "and the flyout at its right"
    );
    assert_eq!(
        split.popup.panel.left,
        split.menu.panel.right + 4,
        "the narrowed flyout is still right of the narrowed menu"
    );
    assert!(split.menu.panel.right < split.popup.panel.left);

    // And the same narrowing at a display's own larger scale.
    let split = menu_flyout_from_menu(
        &menu_popup_from_button(the_gear(), 250, 300, 144, &menu_rows),
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
        &menu_popup_from_button(the_gear(), 600, 120, 96, &menu_rows),
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
        (600i32, 300i32, 96u32), // the flyout at the menu's side
        (400, 300, 96),          // the menu pulled left
        (250, 300, 96),          // the two narrowed
        (500, 300, 144),         // pulled left at a larger scale
        (250, 300, 144),         // narrowed at a larger scale
    ];

    for (width, height, dpi) in placements {
        let placed = menu_flyout_from_menu(
            &menu_popup_from_button(the_gear(), width, height, dpi, &menu_rows),
            width,
            height,
            dpi,
            &flyout_rows,
        );
        let menu = placed.menu;
        let flyout = placed.popup;

        // The flyout is beside the menu, not over it: its panel
        // starts at or after the menu's own right edge, so no
        // row of the menu is covered.
        assert!(
            flyout.panel.left >= menu.panel.right,
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
    let popup = menu_popup_from_button(the_gear(), 400, 300, 96, &three_rows());

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
/// from, and each kind of mark is its own art there: a check row
/// a square, checked where the setting is on and empty where it
/// is off, and a choice of a seek page its disc. The row that
/// carries no mark carries nothing there at all.
#[test]
fn each_kind_of_mark_is_its_own_art_in_the_room_before_the_labels() {
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
    let popup = menu_popup_from_button(the_gear(), width, height, 96, &rows);
    paint_menu_popup(&mut buffer, width, &palette, &popup, &rows, &surface, 1.0);

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
    let far_from = |ink: &[(i32, i32)], at: (f32, f32), past: f32| {
        ink.iter().any(|&(x, y)| {
            let dx = x as f32 + 0.5 - at.0;
            let dy = y as f32 + 0.5 - at.1;
            (dx * dx + dy * dy).sqrt() > past
        })
    };
    let near_to = |ink: &[(i32, i32)], at: (f32, f32), within: f32| {
        ink.iter().all(|&(x, y)| {
            let dx = x as f32 + 0.5 - at.0;
            let dy = y as f32 + 0.5 - at.1;
            (dx * dx + dy * dy).sqrt() <= within
        })
    };

    // The checked row: a square in the room, and the check
    // drawn in it — which is ink the empty square of the row
    // below does not carry.
    let checked = row_ink(0);
    let empty = row_ink(1);
    assert!(!checked.is_empty(), "a checked row carries a box");
    assert!(!empty.is_empty(), "an unchecked row carries an empty box");
    assert!(
        checked.len() > empty.len(),
        "the check is drawn in: the checked box carries more ink than the empty one"
    );

    // The box of a check row is a square rather than a dot:
    // its ink reaches the corners of the room it stands in,
    // far from the mark's own centre.
    assert!(
        far_from(&checked, centre(0), 4.0),
        "the box of a check row is a square, not a dot"
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
    let popup = menu_popup_from_button(the_gear(), width, height, 96, &rows);
    paint_menu_popup(&mut buffer, width, &palette, &popup, &rows, &surface, 1.0);

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
