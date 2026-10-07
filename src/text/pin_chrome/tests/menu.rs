use super::*;

/// The rows the card's menu carries: the two that turn a mode on and
/// off, and the one that opens the seek choices. Which of them is the
/// one in force is what the marks say.
fn three_rows() -> Vec<MenuRow> {
    vec![
        MenuRow {
            label: "Shuffle Mode".to_string(),
            marked: true,
        },
        MenuRow {
            label: "Loop".to_string(),
            marked: false,
        },
        MenuRow {
            label: "Seek".to_string(),
            marked: false,
        },
    ]
}

/// The menu opens downward from the bullet's own row, with its left
/// edge on the bullet's own left edge — the panel hangs from the cell
/// the press came from, the way the volume popup hangs from the button
/// that opened it.
#[test]
fn the_menu_opens_down_from_the_bullet_with_its_left_edge_on_it() {
    let bullet = RECT {
        left: 15,
        top: 12,
        right: 33,
        bottom: 35,
    };
    let rows = three_rows();
    let popup = menu_popup_from_bullet(bullet, 400, 300, 96, &rows);

    assert_eq!(popup.panel.left, bullet.left);
    assert!(
        popup.panel.top >= bullet.bottom,
        "the panel hangs below the bullet's row, not over it"
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

/// A menu that would run off the window is held inside it: a panel
/// off the side of a window is a row a hand cannot reach, and one off
/// the bottom is rows that cannot be read at all.
#[test]
fn a_menu_that_would_run_off_the_window_stays_inside_it() {
    let rows = three_rows();

    // A narrow window: the panel is wider than the room the bullet's
    // own edge leaves to the right, so the panel is held at the
    // window's own left edge rather than running off its right.
    let narrow = menu_popup_from_bullet(
        RECT {
            left: 20,
            top: 4,
            right: 38,
            bottom: 27,
        },
        100,
        300,
        96,
        &rows,
    );
    assert_eq!(narrow.panel.left, 0, "a narrow window holds the panel at its left");
    assert!(narrow.panel.right <= 100);

    // A short window: the bullet sits near the bottom and the panel
    // cannot hang below it inside the window, so the panel is held at
    // the window's own bottom instead.
    let short = menu_popup_from_bullet(
        RECT {
            left: 15,
            top: 90,
            right: 33,
            bottom: 113,
        },
        400,
        120,
        96,
        &rows,
    );
    assert!(short.panel.bottom <= 120, "the panel does not run off the bottom");
    assert!(short.panel.top >= 0, "and it is still a panel of the window");
}

/// A point answers a row only inside the rows: the pad the panel is
/// drawn with above and below them is no row at all, and neither is
/// anywhere outside the panel, above it or below it — a press there is
/// a press outside the menu, which is what closes it.
#[test]
fn a_point_is_a_row_only_inside_the_rows() {
    let bullet = RECT {
        left: 15,
        top: 12,
        right: 33,
        bottom: 35,
    };
    let rows = three_rows();
    let popup = menu_popup_from_bullet(bullet, 400, 300, 96, &rows);

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
    let popup = menu_popup_from_bullet(
        RECT {
            left: 15,
            top: 12,
            right: 33,
            bottom: 35,
        },
        width,
        height,
        96,
        &rows,
    );
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
