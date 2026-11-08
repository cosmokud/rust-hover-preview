use super::*;
// The two halves of the caption's own width rule, which `pin_chrome` re-exports only the second of:
// a name's room and whether there is one are questions the caption asks itself, and the test of that
// rule has to reach them where they live rather than through the glob — which finds the *test* module
// named `caption` first.
use crate::text::pin_chrome::caption::{
    buttons_fit, title_fits, title_fits_beside_pill, title_room,
};

#[test]
fn a_cut_name_is_the_whole_name_where_it_fits_and_the_longest_beginning_where_it_does_not() {
    // Six pixels a character, so a budget of 24 is four characters and no more. The width is
    // a closure rather than a font so the arithmetic of a name can be pinned down exactly;
    // `a_real_font_cuts_the_name_where_the_measured_width_runs_out` is what pins the font.
    let mut width = |text: &str| text.chars().count() as i32 * 6;

    assert_eq!(longest_prefix_that_fits("abcdef", 24, &mut width), "abcd…");
    assert_eq!(longest_prefix_that_fits("abcdef", 30, &mut width), "abcde…");
    // Nothing fits beside the ellipsis, so nothing is drawn — a lone ellipsis is not a name.
    assert_eq!(longest_prefix_that_fits("abcdef", 0, &mut width), "");
    assert_eq!(longest_prefix_that_fits("abcdef", 5, &mut width), "");
    // One boundary that fits, and the name is kept to its first character.
    assert_eq!(longest_prefix_that_fits("abcdef", 6, &mut width), "a…");
    // A single character has no boundary to cut at, so it is not cut either.
    assert_eq!(longest_prefix_that_fits("a", 0, &mut width), "");
}

#[test]
fn a_cut_name_is_cut_on_a_character_rather_than_inside_one() {
    // A name of multi-byte characters: every boundary the search tries is the start of a
    // character, so a cut cannot land in the middle of one and leave half a glyph.
    let title = "äöüäöüäöü";
    let mut width = |text: &str| text.chars().count() as i32 * 10;

    let cut = longest_prefix_that_fits(title, 25, &mut width);

    assert_eq!(cut, "äö…");
    assert!(title.starts_with(cut.trim_end_matches('…')));
    assert!(cut.is_char_boundary(cut.len()));
}

/// The search and the walk it replaced must answer the same thing for every budget, because
/// the search is the walk with fewer measurements and the whole claim of the change is that
/// the caption is not shorter for it.
///
/// The property is asked of `longest_prefix_that_fits`, which is the function `fit_title`
/// draws its name through — a second copy of the search would agree with this loop whatever
/// the caption did, which is how the three tests above it came to pass with `fit_title`
/// deleted outright.
#[test]
fn the_search_agrees_with_walking_it() {
    let title = "D:\\some\\folder\\a rather long file name.txt";
    let width = |text: &str| text.chars().count() as i32 * 7;

    for budget in 0..(title.chars().count() as i32 * 7) {
        let boundaries: Vec<usize> = title
            .char_indices()
            .skip(1)
            .map(|(index, _)| index)
            .collect();

        let mut walked = 0usize;
        for &index in &boundaries {
            if width(&title[..index]) > budget {
                break;
            }
            walked = index;
        }

        let expected = if walked == 0 {
            String::new()
        } else {
            format!("{}…", &title[..walked])
        };

        assert_eq!(
            longest_prefix_that_fits(title, budget, &mut |text| width(text)),
            expected,
            "the search and the walk differ at a budget of {budget}"
        );
    }
}

/// A caption is cut by what its own font measures, and not by anything a test can arrange:
/// every width above is six or ten pixels a character because a real one is not, and a name
/// cut at the wrong place in a real font is a name with its last character gone.
///
/// So the search is asked through `fit_title` on a surface with a font in it — the path
/// that has no test at all when the search is copied rather than shared, which is what left
/// this function free to be wrong for as long as it was here.
#[test]
fn a_real_font_cuts_the_name_where_the_measured_width_runs_out() {
    let palette = ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    };
    let surface = DibSurface::create(600, 30).expect("a surface");
    let style = caption_style(palette.foreground);
    let title = "D:\\some\\folder\\a rather long file name.txt";

    // The whole name where it fits, exactly, and something cut a character or two off where
    // it does not — measured rather than counted, so the last few characters go in the order
    // their own widths say they can and not in the order a name reads.
    let whole = measure_text(&surface, &style, title, 1.0);
    assert_eq!(fit_title(&surface, &style, title, whole, 1.0), title);

    let cut = fit_title(&surface, &style, title, whole - 1, 1.0);
    assert!(
        title.starts_with(cut.strip_suffix('…').unwrap_or_default()),
        "`{cut}` is not a beginning of `{title}`"
    );
    assert!(
        title.len() - cut.trim_end_matches('…').len() > 1,
        "a name one pixel too wide loses the ellipsis and at least one character to it"
    );

    // And nothing at all where even the first character will not fit beside the ellipsis,
    // which a search over a made-up width could not have shown.
    assert_eq!(fit_title(&surface, &style, title, 1, 1.0), "");
    assert_eq!(fit_title(&surface, &style, title, 0, 1.0), "");

    // Every room in between: what comes back is a beginning of the name with an ellipsis
    // after it, it is not the whole name, and it is as wide as the room allows — never
    // wider, which is the whole claim of measuring rather than counting.
    for available in 1..whole {
        let fitted = fit_title(&surface, &style, title, available, 1.0);
        assert!(
            title.starts_with(fitted.strip_suffix('…').unwrap_or(&fitted)),
            "`{fitted}` is not a beginning of `{title}`"
        );
        assert_ne!(fitted, title, "a name that does not fit was not cut");
        assert!(
            measure_text(&surface, &style, &fitted, 1.0) <= available,
            "`{fitted}` is wider than the {available} pixels it was given"
        );
    }
}

/// Every button a caption carries is found by pointing at it. Painting and hit-testing are
/// both read off the one list of boxes, so a button that is drawn and is not on, or is on
/// and is not drawn, is a caption the hand and the eye disagree about.
#[test]
fn every_button_of_a_caption_is_the_one_under_the_pointer() {
    let buttons = button_boxes(600, 30, 96, true);
    assert_eq!(buttons.len(), 9);

    for button in &buttons {
        let middle_x = (button.rect.left + button.rect.right) / 2;
        assert_eq!(
            button_at(middle_x, 5, 600, 30, 96, true),
            Some(button.kind),
            "the middle of {:?} is not on it",
            button.kind
        );
        // Each end of the box is its own, and the pixel past the last one is the title's.
        assert_eq!(
            button_at(button.rect.left, 0, 600, 30, 96, true),
            Some(button.kind)
        );
        assert_eq!(
            button_at(button.rect.right - 1, 29, 600, 30, 96, true),
            Some(button.kind)
        );
    }

    assert_eq!(
        button_at(buttons[0].rect.left - 1, 5, 600, 30, 96, true),
        None,
        "the strip to the left of the walk is the title's"
    );
}

/// A caption too narrow for the walk carries the window's buttons and nothing else. A pin is worth
/// less than a window it can be closed from, and the buttons that stay are the ones a hand has been
/// reaching for on every window it has ever had — so the walk goes first, from its own left end, and
/// close keeps the right edge of the strip.
#[test]
fn a_caption_narrower_than_its_own_walk_keeps_the_window_s_buttons() {
    let narrow = button_boxes(200, 30, 96, true);

    // Buttons are 46 wide and never narrower, so a strip 200 wide holds four of them: the window's
    // own, taken from close at the right edge back to the left (see `button_boxes`) — the detach
    // among them, because the road out of being a pin is a window's road and stands with them.
    assert_eq!(narrow.len(), 4);
    assert_eq!(narrow[0].kind, CaptionButton::Detach);
    assert_eq!(narrow[1].kind, CaptionButton::Minimize);
    assert_eq!(narrow[2].kind, CaptionButton::Maximize);
    assert_eq!(narrow[3].kind, CaptionButton::Close);
    assert_eq!(narrow[3].rect.right, 200);

    // And the pointer agrees with the painter about what is there.
    for button in &narrow {
        assert_eq!(
            button_at(button.rect.left + 1, 5, 200, 30, 96, true),
            Some(button.kind)
        );
    }
    assert_eq!(button_at(0, 5, 200, 30, 96, true), None);
    for kind in [
        CaptionButton::Previous,
        CaptionButton::Next,
        CaptionButton::OpenWith,
        CaptionButton::OpenWithList,
    ] {
        assert!(
            narrow.iter().all(|button| button.kind != kind),
            "a caption too narrow for the walk does not carry {kind:?}"
        );
    }
}

/// A caption narrows one button at a time, and in the one order a window owes: the walk from its own
/// left end, then the window's own from nothing to close last, so that the button that closes the
/// window is the one a caption keeps longest.
///
/// The widths are the boundaries of the rule rather than a sample of it: each one is either side of a
/// button being lost, because a caption that dropped two at once, or the wrong one, would still answer
/// right at every width in the middle of a step.
#[test]
fn a_caption_gives_up_its_buttons_one_at_a_time_from_the_left() {
    const ORDER: [CaptionButton; 9] = [
        CaptionButton::Previous,
        CaptionButton::Next,
        CaptionButton::OpenWith,
        CaptionButton::OpenWithList,
        CaptionButton::AlwaysOnTop,
        CaptionButton::Detach,
        CaptionButton::Minimize,
        CaptionButton::Maximize,
        CaptionButton::Close,
    ];

    let boundaries = [
        (9 * 46, 9),
        (9 * 46 - 1, 8),
        (8 * 46, 8),
        (8 * 46 - 1, 7),
        (7 * 46, 7),
        (7 * 46 - 1, 6),
        (6 * 46, 6),
        (6 * 46 - 1, 5),
        (5 * 46, 5),
        (5 * 46 - 1, 4),
        (4 * 46, 4),
        (4 * 46 - 1, 3),
        (3 * 46, 3),
        (3 * 46 - 1, 2),
        (2 * 46, 2),
        (2 * 46 - 1, 1),
        (46, 1),
        (45, 0),
        (1, 0),
    ];
    for (width, kept) in boundaries {
        let boxes = button_boxes(width, 30, 96, true);
        assert_eq!(
            boxes.iter().map(|button| button.kind).collect::<Vec<_>>(),
            ORDER[ORDER.len() - kept..],
            "a caption {width} wide carries {kept} buttons, taken from the right"
        );
        assert_eq!(buttons_fit(width, 96, true), kept == ORDER.len());

        // Close is against the right edge and its neighbours stand beside it without a gap: the
        // survivors keep the places they had on a caption wide enough to carry all of them.
        assert_eq!(
            boxes.last().map(|button| (button.kind, button.rect.right)),
            (kept > 0).then_some((CaptionButton::Close, width)),
            "close is the last button and the right edge of the strip, at width {width}"
        );
        for pair in boxes.windows(2) {
            assert_eq!(pair[0].rect.right, pair[1].rect.left, "no gap at {width}");
            assert_eq!(
                pair[1].rect.right - pair[1].rect.left,
                46,
                "no shrinking at {width}"
            );
        }
    }
}

/// A caption too narrow for even close is still a bar: the background with its hairline under it, and
/// nothing on it to press. A title bar that vanished when the window was dragged narrow and came back
/// on the way out would be a band that flickered with the drag; this one just runs out of buttons.
#[test]
fn a_caption_too_narrow_for_close_still_paints_its_bar() {
    let palette = ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    };
    let (width, height) = (30u32, 30u32);
    assert!(
        button_boxes(width as i32, height as i32, 96, true).is_empty(),
        "a caption narrower than one button has none"
    );

    let surface = DibSurface::create(width, height).expect("a surface");
    let caption = Caption {
        title: "notes.txt",
        read_only: false,
        maximized: false,
        maximizable: true,
        always_on_top: true,
        hovered: None,
        pressed: None,
    };
    paint_caption(&surface, &palette, &caption, 96);

    // The surface is the desktop's own byte order, so a pixel is read back the way it was written.
    let pixels = surface.pixels();
    let at = |x: u32, y: u32| {
        let index = (y as usize * width as usize + x as usize) * 4;
        [pixels[index + 2], pixels[index + 1], pixels[index]]
    };
    let hairline = palette.hover(0.18);
    for y in 0..height - 1 {
        for x in 0..width {
            assert_eq!(
                at(x, y),
                palette.background,
                "the bar is its own background at {x},{y}"
            );
        }
    }
    for x in 0..width {
        assert_eq!(
            at(x, height - 1),
            hairline,
            "and its hairline runs the whole of the bottom row, at {x}"
        );
    }
}

/// A name is not said beside a row of buttons a caption has given up, even where the room the missing
/// ones left would be room for a name: the name belongs to the caption that is whole, and one said
/// beside a row that grows and shrinks as the window is dragged would be a title whose own buttons
/// moved out from under a hand reaching for one (see `title_fits`).
#[test]
fn a_name_is_not_said_beside_buttons_a_caption_has_given_up() {
    let palette = ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    };
    let strip = |width: u32, title: &str| {
        let surface = DibSurface::create(width, 30).expect("a surface");
        let caption = Caption {
            title,
            read_only: false,
            maximized: false,
            maximizable: true,
            always_on_top: true,
            hovered: None,
            pressed: None,
        };
        paint_caption(&surface, &palette, &caption, 96);
        surface.pixels()
    };

    // A caption 240 wide has dropped four of its nine buttons and still says nothing.
    assert!(
        !buttons_fit(240, 96, true) && !title_fits(240, 96, true),
        "a caption with buttons missing has given the name up"
    );
    assert_eq!(
        strip(240, "notes.txt"),
        strip(240, ""),
        "a strip with buttons missing painted a name beside them"
    );

    // And a caption that carries all nine says the name it has room for.
    assert!(
        buttons_fit(520, 96, true) && title_fits(520, 96, true),
        "a caption with every button has room for a name beside them"
    );
    assert_ne!(
        strip(520, "notes.txt"),
        strip(520, ""),
        "a whole caption did not paint the name it had room for"
    );

    // The two questions cannot disagree: no width says a name beside a row that is short a button.
    for width in 1..1200 {
        assert!(
            !(title_fits(width, 96, true) && !buttons_fit(width, 96, true)),
            "a caption {width} wide gave the name room beside buttons it had given up"
        );
    }
}

/// A pin with nothing to maximize — a sound's card — carries the two buttons that mean
/// something on it, packed against the right edge the way Windows packs a window's: the one
/// that closes it keeps the place it has on every other caption, and the one beside it is
/// the minimize that was there before. Neither of them moves for the walk, which sits
/// against the group and is measured off the same strip.
#[test]
fn a_caption_without_a_maximize_keeps_the_close_button_where_it_was() {
    let three = button_boxes(600, 30, 96, true);
    let two = button_boxes(600, 30, 96, false);

    assert_eq!(two.len(), 8);
    assert_eq!(two[0].kind, CaptionButton::Previous);
    assert_eq!(two[1].kind, CaptionButton::Next);
    assert_eq!(two[2].kind, CaptionButton::OpenWith);
    assert_eq!(two[3].kind, CaptionButton::OpenWithList);
    assert_eq!(two[4].kind, CaptionButton::AlwaysOnTop);
    assert_eq!(two[5].kind, CaptionButton::Detach);
    assert_eq!(two[6].kind, CaptionButton::Minimize);
    assert_eq!(two[7].kind, CaptionButton::Close);

    let close = three
        .iter()
        .find(|button| button.kind == CaptionButton::Close)
        .expect("a close button")
        .rect;
    let minimize = three
        .iter()
        .find(|button| button.kind == CaptionButton::Minimize)
        .expect("a minimize button")
        .rect;

    assert_eq!(two[7].rect.left, close.left);
    assert_eq!(two[7].rect.right, close.right);
    assert_eq!(two[6].rect.right, two[7].rect.left);
    assert_eq!(two[5].rect.right, two[6].rect.left);
    assert_eq!(
        two[6].rect.right - two[6].rect.left,
        minimize.right - minimize.left
    );

    // And the space the button used to take is a button's, not a hole: it is the minimize
    // that has moved along into it, with the walk along beside it.
    let maximize = three
        .iter()
        .find(|button| button.kind == CaptionButton::Maximize)
        .expect("a maximize button")
        .rect;
    assert_eq!(
        button_at(maximize.left + 1, 5, 600, 30, 96, false),
        Some(CaptionButton::Minimize)
    );
    assert_eq!(
        button_at(maximize.left + 1, 5, 600, 30, 96, true),
        Some(CaptionButton::Maximize)
    );
}

/// A caption is opaque everywhere, and a tooltip drawn on one of its own surfaces carries
/// its text across with the coverage that text has.
///
/// The name is written through GDI, and GDI knows nothing of an alpha channel: it leaves the
/// alpha byte of everything it draws at zero. A caption is copied into the layered surface
/// a row at a time and composited with `AC_SRC_ALPHA`, so a pixel left at zero is a hole in
/// the title bar — and a run of text is drawn with an opaque background over the whole of
/// its box, not just over its glyphs, so the hole is a box rather than a letter.
///
/// This is the one thing about a caption's text a test on its layout cannot see, and the
/// reason the caption hands its own coverage back after the last of its runs.
#[test]
fn a_caption_is_opaque_everywhere_it_is_painted() {
    let palette = ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    };
    let width = 600u32;
    // A caption is as tall as a pinned window's caption is given at a display's scale, which
    // is the height the surface is really created at (see `pinned_caption_height`).
    let height = 30u32;
    let surface = DibSurface::create(width, height).expect("a surface");

    // With a name on it and without: the coverage is the caption's own, and a caption that
    // only happened to be opaque on a strip with nothing written on it is not opaque.
    for title in ["picture.png", ""] {
        let caption = Caption {
            title,
            read_only: false,
            maximized: false,
            maximizable: true,
            always_on_top: true,
            hovered: Some(CaptionButton::OpenWith),
            pressed: None,
        };
        paint_caption(&surface, &palette, &caption, 96);

        let pixels = unsafe {
            std::slice::from_raw_parts(surface.bits(), width as usize * height as usize * 4)
        };
        let transparent = pixels
            .as_chunks::<4>()
            .0
            .iter()
            .filter(|pixel| pixel[3] != 255)
            .count();
        assert_eq!(
            transparent, 0,
            "a caption is opaque wherever it is painted, and {transparent} of its pixels were not"
        );
    }
}

/// The strip a repaint copies is the strip that was painted, and every change that would
/// have moved a pixel on it is a change that misses the cache.
///
/// A cached caption that does not repaint is the worst failure this module has available: a
/// window whose title bar shows the last file's name, or a button that stays washed under a
/// pointer that has walked off it, with nothing in the window to say so. So each input the
/// drawing reads is changed on its own and the strip is required to change with it — and
/// then the *same* strip is asked for twice and required to come back byte for byte, which
/// is what makes the cache worth having at all.
#[test]
fn a_caption_is_repainted_when_anything_it_is_drawn_from_changes() {
    let palette = ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    };
    let dark = ChromePalette {
        background: [30, 30, 30],
        foreground: [220, 220, 220],
        accent: [10, 140, 240],
        dark: true,
    };
    // A name short enough to be drawn whole on the strip: a caption whose name is cut away
    // entirely is the same picture whichever file it is, so a test on one of those could not
    // tell a stale strip from a changed one.
    let caption = Caption {
        title: "Readme.md",
        read_only: false,
        maximized: false,
        maximizable: true,
        always_on_top: true,
        hovered: Some(CaptionButton::OpenWith),
        pressed: None,
    };
    let surface = DibSurface::create(600, 30).expect("a surface");

    let strip = |palette: &ChromePalette, caption: &Caption, dpi: u32| {
        paint_caption(&surface, palette, caption, dpi);
        surface.pixels()
    };

    let first = strip(&palette, &caption, 96);
    // The same strip again, and byte for byte the one before: a cached caption is the bytes
    // the eye has already been shown, or it is a caption that flickers for nothing.
    assert_eq!(
        strip(&palette, &caption, 96),
        first,
        "the same caption asked for twice is not the strip it was"
    );

    // And every one of these is a different strip. Each is asked for straight after the
    // first one rather than after the one before it, so that only what it changes differs —
    // a cache missing one field is only caught by a change that is *only* that field.
    let changed = [
        (
            "another file",
            &palette,
            &Caption {
                title: "Cargo.toml",
                ..caption
            },
            96,
        ),
        (
            "the hovered button left",
            &palette,
            &Caption {
                hovered: None,
                ..caption
            },
            96,
        ),
        (
            "the pressed button arrived",
            &palette,
            &Caption {
                hovered: None,
                pressed: Some(CaptionButton::Close),
                ..caption
            },
            96,
        ),
        (
            "a close button is washed red and another one is not",
            &palette,
            &Caption {
                hovered: Some(CaptionButton::Close),
                ..caption
            },
            96,
        ),
        (
            "a window with a maximum to go to and one without",
            &palette,
            &Caption {
                maximizable: false,
                ..caption
            },
            96,
        ),
        (
            "a maximized window draws a restore pair and an ordinary one does not",
            &palette,
            &Caption {
                maximized: true,
                ..caption
            },
            96,
        ),
        (
            "always on top toggled",
            &palette,
            &Caption {
                always_on_top: false,
                ..caption
            },
            96,
        ),
        ("another display's scale", &palette, &caption, 144),
        ("another theme", &dark, &caption, 96),
    ];
    for (what, other_palette, other_caption, other_dpi) in changed {
        // Straight back to the strip painted above, so the cache holds it and the only thing
        // that differs about the next paint is what this case changes.
        assert_eq!(
            strip(&palette, &caption, 96),
            first,
            "{what}: the base moved"
        );
        assert_ne!(
            strip(other_palette, other_caption, other_dpi),
            first,
            "`{what}` painted the caption that was already there"
        );
    }

    // And a wider window is not a narrow one's strip with room at the end of it: the buttons
    // are against the right edge and the name is cut against the left one, so both of them
    // move with the window.
    let wide = DibSurface::create(900, 30).expect("a surface");
    paint_caption(&wide, &palette, &caption, 96);
    assert_ne!(
        wide.pixels(),
        first,
        "a resized window painted the strip of the width it had"
    );
}

/// The pictures this test writes are the design under review: the marks the caption's own
/// buttons carry, at both scales, side by side in one strip. A row of marks is judged as a
/// row — one of them twice the size of the rest, or a third the size, or sitting a pixel or
/// two off the middle of its own button, is what a strip shows and a test on one mark at a
/// time cannot.
#[test]
fn draws_the_caption_glyphs() {
    let dir = scratch("caption-glyphs");
    // The marks on the bar's own background: the glyphs are ink, so a picture that is not
    // the colour behind them shows nothing, which is what a black buffer and black ink do.
    let bar = [250u8, 250, 250];
    let glyph_ink = [56u8, 58, 66];
    let marks = [
        ("previous", CaptionButton::Previous, false),
        ("next", CaptionButton::Next, false),
        ("open-with", CaptionButton::OpenWith, false),
        ("open-with-list", CaptionButton::OpenWithList, false),
        ("always-on-top", CaptionButton::AlwaysOnTop, false),
        ("detach", CaptionButton::Detach, false),
        ("minimize", CaptionButton::Minimize, false),
        ("maximize", CaptionButton::Maximize, false),
        ("restore", CaptionButton::Maximize, true),
        ("close", CaptionButton::Close, false),
    ];

    for (dpi, tag) in [(96u32, "96dpi"), (192u32, "192dpi")] {
        let scale = dpi as f32 / 96.0;
        let side = text_paint::scaled(BUTTON_PIXELS as i32, scale).max(1);
        let height = side as u32;
        let width = (side * marks.len() as i32) as u32;
        let mut out = image_buffer(width, height);
        let mut sizes: Vec<(&str, i32, i32, i32, i32)> = Vec::new();

        for (index, (name, kind, maximized)) in marks.iter().enumerate() {
            let surface = DibSurface::create(side as u32, height).expect("a surface");
            // The bar behind the mark, opaque, so that a column the glyph did not reach
            // reads as the background and not as ink. Filled through the same raw bits
            // every glyph here is drawn into.
            let filled = unsafe {
                std::slice::from_raw_parts_mut(surface.bits(), 4 * side as usize * height as usize)
            };
            for pixel in filled.as_chunks_mut::<4>().0 {
                pixel[0] = bar[2];
                pixel[1] = bar[1];
                pixel[2] = bar[0];
                pixel[3] = 255;
            }

            paint_glyph(
                &surface,
                CaptionButtonBox {
                    kind: *kind,
                    rect: RECT {
                        left: 0,
                        top: 0,
                        right: side,
                        bottom: side,
                    },
                },
                *maximized,
                glyph_ink,
                scale,
            );

            let painted = surface.pixels();
            let ink = |x: usize, y: usize| painted[(y * side as usize + x) * 4] < 128;

            // The external link is drawn in whole pixels, not stroked: a stroked diagonal
            // leaves a ramp of blended pixels down every row it crosses, and the mark is
            // asked to be sharp. Sharp here means every pixel is either the bar or the ink,
            // with nothing in between.
            if *name == "open-with" {
                for (index, pixel) in painted.as_chunks::<4>().0.iter().enumerate() {
                    assert!(
                        pixel[0] == bar[2] || pixel[0] == glyph_ink[2],
                        "{tag} the external link is not sharp at pixel {index}: {}",
                        pixel[0]
                    );
                }

                // And no pixel of it is an island. A one-pixel diagonal touches its own
                // neighbours only at their corners, which is the dotted staircase the shaft
                // is drawn a pixel thicker to avoid; every pixel of this mark has to share
                // an edge with another.
                for y in 0..height as usize {
                    for x in 0..side as usize {
                        if !ink(x, y) {
                            continue;
                        }
                        let beside = (x > 0 && ink(x - 1, y))
                            || (x + 1 < side as usize && ink(x + 1, y))
                            || (y > 0 && ink(x, y - 1))
                            || (y + 1 < height as usize && ink(x, y + 1));
                        assert!(
                            beside,
                            "{tag} the external link has an island pixel at ({x}, {y})"
                        );
                    }
                }
            }

            // The mark's own bounds in the surface it was painted into, as the min and
            // max of every row and column that carries any ink at all.
            let mut low = (i32::MAX, i32::MAX);
            let mut high = (i32::MIN, i32::MIN);
            for y in 0..height as usize {
                for x in 0..side as usize {
                    if !ink(x, y) {
                        continue;
                    }
                    low = (low.0.min(x as i32), low.1.min(y as i32));
                    high = (high.0.max(x as i32), high.1.max(y as i32));
                }
            }
            assert!(
                low.0 > i32::MIN && high.0 > i32::MIN,
                "{name} drew nothing at all at {tag}"
            );

            // The mark sits in the middle of its own button, which is the half of "beauty"
            // a row of glyphs is judged on that no single-glyph test can see: a mark one
            // pixel off its centre is a mark the hand does not aim where the eye is.
            let middle = side / 2;
            let (across, down) = ((low.0 + high.0) / 2 - middle, (low.1 + high.1) / 2 - middle);
            assert!(
                across.abs() <= 1 && down.abs() <= 1,
                "{tag} {name} is {across}px across and {down}px down from the middle of its button"
            );

            // And it is the size of the rest of the row, which is the other half. A glyph
            // drawn twice the height of its neighbours is the one the report is about, and
            // a picture is the only place that shows: the chevron once reached a whole span
            // either side of the middle row, twice every other mark's height. The
            // minimize bar is a bar, so it is not held to the height of the rest.
            let (tall, wide) = (high.1 - low.1 + 1, high.0 - low.0 + 1);
            sizes.push((*name, tall, wide, low.1, high.1));

            println!(
                "{tag:>6} {name:<9} x {}..={} y {}..={} ({}x{})",
                low.0,
                high.0,
                low.1,
                high.1,
                high.0 - low.0 + 1,
                high.1 - low.1 + 1
            );

            let at = (index as u32) * side as u32;
            blit_cell(&mut out, width, height, &painted, side as u32, at);
        }

        // The walk buttons are the size of the marks they sit beside, which is the whole
        // of the report they were redrawn for: they used to reach a whole span either side
        // of the middle row and come out twice the height of every other mark in the bar.
        //
        // The size they have to match is the glyph square the strip is drawn on — a mark
        // is asked for a span, and a span across and a span down is what it should be. The
        // close cross is not the yardstick: a diagonal stroke laid down a pixel at a time
        // overshoots its own ends by the stroke's width, so the cross is a row or two taller
        // than it is wide at 200% and always has been. Comparing to it would be holding the
        // chevron to the cross's overhang rather than to the size a glyph is asked for.
        let glyph = text_paint::scaled(GLYPH_PIXELS as i32, scale).max(6);
        let square = glyph + 1;
        for (name, tall, wide, ..) in &sizes {
            if *name != "previous" && *name != "next" {
                continue;
            }
            assert_eq!(
                (*tall, *wide),
                (square, square),
                "{tag} {name} is {tall}x{wide} beside the {square}x{square} a glyph is asked for"
            );
        }

        // And the row is one size the rest of the way: nothing stands much taller than the
        // rest, which is the property a picture shows and a per-glyph test cannot. The
        // minimize bar is not counted: it is a bar, and is one row tall by being a bar.
        let heights = sizes
            .iter()
            .filter(|(name, ..)| *name != "minimize")
            .map(|(_, tall, ..)| *tall);
        let tallest = heights.clone().max().unwrap_or_default();
        let smallest = heights.min().unwrap_or_default();
        assert!(
            tallest - smallest <= 2,
            "{tag} the row runs from {smallest} to {tallest} rows tall"
        );

        // The always-on-top pin and the external-link mark are the height of the window's own
        // marks beside them, which is what a hand reads before it reads any of their shapes:
        // the pin's ceiling and the link's corner on the row a maximize's top edge is on, and
        // the pin's stem and the link's box bottom on the row that box's bottom edge is on. A
        // mark drawn to a smaller square inside its button is the one a row of them shows as
        // the odd one out, whatever its own lines are drawn like.
        let rows = |name: &str| {
            let (_, _, _, top, bottom) = sizes
                .iter()
                .find(|(each, ..)| *each == name)
                .unwrap_or_else(|| panic!("{tag} has no {name}"));
            (*top, *bottom)
        };
        for mark in ["always-on-top", "open-with"] {
            let mark_rows = rows(mark);
            for beside in ["maximize", "close"] {
                assert_eq!(
                    mark_rows,
                    rows(beside),
                    "{tag} {mark} runs {mark_rows:?} beside {beside}'s {:?}",
                    rows(beside)
                );
            }
        }

        let path = dir.join(format!("caption-glyphs-{tag}.png"));
        write_png(path.clone(), &out, width, height);
        println!("{}", path.display());
    }
}

/// The two walk buttons point opposite ways, each points the way it walks, and each is
/// drawn the size of every other glyph in the bar.
///
/// A chevron whose arms open on both sides of its middle is a cross, and one whose arms
/// open away from the end the columns stop at points the wrong way — either of which is a
/// button that reads as something other than what it does, and neither of which a test on
/// the boxes alone could see. Arms that reach further than the mark is wide are the third
/// of the same kind: a row of marks that are not one size is a row of things that are not
/// one kind of thing, and the walk buttons are the two biggest of the six.
///
/// The ink is measured by coverage here and not by whole pixels, which is the measure the
/// glyph dump above uses and the one the assertions are about: the arms are anti-aliased, and
/// a band a pixel wide down a slope covers no pixel whole — nine tenths of one is the most
/// there is at this slope — so counting whole pixels would be asking this test for the walk
/// the arms used to be drawn with rather than for the shape a chevron has.
#[test]
fn the_two_walk_buttons_are_chevrons_pointing_opposite_ways() {
    // A box with room for the whole glyph and a row to spare: the arms part half a span
    // either side of the middle row, so a box the glyph's own width would cut the open
    // end off and leave a test that passes for a chevron that is really a stub.
    const W: i32 = 24;
    const H: i32 = 24;
    let span = 10;
    let ink = [0u8, 0, 0];

    // The rows carrying ink in one column. The buffer is filled with a colour first, so that
    // a blank column reads as blank rather than as ink, and a pixel counts as ink where it is
    // more ink than bar — which is more than half of it, the ink being black and the bar white.
    fn drawn(buffer: &[u8], x: i32) -> Vec<i32> {
        (0..H)
            .filter(|y| buffer[((*y * W + x) * 4) as usize] < 128)
            .collect()
    }

    let mut left = vec![255u8; (W * H * 4) as usize];
    draw_chevron(&mut left, W, 12, 12, span, 1.0, ink, false);
    let mut right = vec![255u8; (W * H * 4) as usize];
    draw_chevron(&mut right, W, 12, 12, span, 1.0, ink, true);

    // A chevron has a vertex: the end it points at is a single row, and the arms open
    // away from it to two that part as they go. A cross has two rows at both ends and
    // four through its middle, so the end that is one row is what says a chevron — and
    // the point is at the end the button walks off, which is a different end for each.
    let point_end = 12 - span / 2;
    let open_end = 12 + span / 2;
    for (buffer, point, open, name) in [
        (&left, point_end, open_end, "the back one"),
        (&right, open_end, point_end, "the on one"),
    ] {
        assert_eq!(
            drawn(buffer, point).len(),
            1,
            "{name} points at one row of its end, and not two: {:?}",
            drawn(buffer, point)
        );
        assert_eq!(
            drawn(buffer, open).len(),
            2,
            "{name} opens to two rows at the other end: {:?}",
            drawn(buffer, open)
        );
    }

    // The two are the same shape facing the other way, so what the back one draws in a
    // column is what the on one draws in the column reflected about the middle of the
    // glyph. Two buttons drawn the same way round would be equal column for column
    // instead, and this is what catches that. Only the columns the glyph itself covers
    // are compared: the rest of the box is empty on both sides and reflects out of it.
    for x in 12 - span / 2..=12 + span / 2 {
        assert_eq!(
            drawn(&left, x),
            drawn(&right, 2 * 12 - x),
            "the two chevrons are each other turned about at x={x}"
        );
    }

    // The arms part half a span either side of the middle row, which is what the close
    // cross and the maximize box are drawn at: a chevron taller than the glyph beside it
    // is an arrow in a row of marks, not a mark in a row.
    for buffer in [&left, &right] {
        let ink = |x: i32, y: i32| buffer[((y * W + x) * 4) as usize] == 0;
        let (top, bottom) = (0..H)
            .find(|y| (point_end..=open_end).any(|x| ink(x, *y)))
            .zip(
                (0..H)
                    .rev()
                    .find(|y| (point_end..=open_end).any(|x| ink(x, *y))),
            )
            .expect("a chevron that was drawn");
        assert_eq!(
            bottom - top + 1,
            span + 1,
            "a chevron is a glyph tall, whatever else it is"
        );
    }
}

/// The two steps of a walk are one another's mirror: a bar at the far edge of the mark and a
/// triangle beside it pointing the way the step goes — ⏮ beside a step back and ⏭ beside a
/// step on, which is what a hand has been reaching for beside a play button for thirty years.
///
/// The bar is the wall the triangle is pushed off, and it is at a different end for each of
/// them; so is the triangle's wide end. A mark whose triangle is walked from the wrong end is
/// a mark whose bar and triangle are the wrong distance apart and, once the walk runs past
/// the far edge of the button, a mark cut in half at one end and drawn twice at the other —
/// none of which a test on the boxes alone can see.
#[test]
fn the_two_steps_of_a_walk_are_mirrors_of_one_another() {
    const W: i32 = 32;
    const H: i32 = 24;
    let centre = 16;
    let span = 14;
    let ink = [0u8, 0, 0];

    // The rows drawn in one column. The buffer is filled with a colour first, so that a
    // blank column reads as blank rather than as ink.
    fn drawn(buffer: &[u8], x: i32) -> Vec<i32> {
        (0..H)
            .filter(|y| buffer[((*y * W + x) * 4) as usize] == 0)
            .collect()
    }

    let mut on = vec![255u8; (W * H * 4) as usize];
    draw_track_step(&mut on, W, centre, 12, span, ink, true);
    let mut back = vec![255u8; (W * H * 4) as usize];
    draw_track_step(&mut back, W, centre, 12, span, ink, false);

    // The two are the same mark facing the other way, so what the back one draws in a column
    // is what the on one draws in the column reflected about the centreline — which is the line
    // between the two middle columns, a whole mark being an even number of columns wide. Two
    // marks drawn the same way round would be equal column for column instead, and a mark drawn
    // a column too wide would fail only on one of the two ends, and this is what catches that.
    for x in 0..W {
        let mirror = 2 * centre - 1 - x;
        if !(0..W).contains(&mirror)
            || (drawn(&back, x).is_empty() && drawn(&on, mirror).is_empty())
        {
            continue;
        }
        assert_eq!(
            drawn(&back, x),
            drawn(&on, mirror),
            "the two steps are each other turned about at x={x}"
        );
    }

    // And both are inside the button they are drawn in, which is what a walk past the far edge
    // would otherwise have lost: neither mark is a column wider than the other, or wider than
    // the button.
    for (buffer, name) in [(&on, "the on one"), (&back, "the back one")] {
        let mut columns = (0..W).filter(|x| !drawn(buffer, *x).is_empty());
        assert_eq!(
            columns.next(),
            Some(centre - span / 2),
            "{name} begins at its own near edge"
        );
        assert_eq!(
            columns.next_back(),
            Some(centre + span / 2 - 1),
            "{name} ends at its own far one"
        );
    }
}

/// The close mark is two diagonals, and a diagonal is the one thing a mark drawn a whole column
/// at a time cannot be: the pixels it crosses are some of the bar and some of the ink, and a
/// painter that gives each of them the one or the other draws a staircase (see `draw_cross`).
#[test]
fn the_close_mark_carries_a_share_of_the_ink_at_its_edges() {
    let side = 24usize;
    // A white bar under a dark ink, so that a pixel which is neither is a share of both: at full
    // coverage the pixel is the ink exactly, and at none of it the bar.
    let bar = [255u8, 255, 255];
    let ink = [56u8, 58, 66] as [u8; 3];

    let surface = DibSurface::create(side as u32, side as u32).expect("a surface");
    let filled = unsafe { std::slice::from_raw_parts_mut(surface.bits(), 4 * side * side) };
    for pixel in filled.as_chunks_mut::<4>().0 {
        pixel[0] = bar[2];
        pixel[1] = bar[1];
        pixel[2] = bar[0];
        pixel[3] = 255;
    }

    paint_glyph(
        &surface,
        CaptionButtonBox {
            kind: CaptionButton::Close,
            rect: RECT {
                left: 0,
                top: 0,
                right: side as i32,
                bottom: side as i32,
            },
        },
        false,
        ink,
        1.0,
    );

    // The two colours a pixel of the mark can be, in the buffer's own order, and the byte that
    // tells them apart: a pixel that is neither is carrying a share of the stroke.
    let painted = surface.pixels();
    let partial = painted
        .as_chunks::<4>()
        .0
        .iter()
        .filter(|pixel| pixel[0] != ink[2] && pixel[0] != bar[2])
        .count();

    assert!(
        partial > 10,
        "the close mark carries part of the ink in {partial} pixels only: the pixels its \
         diagonals cross hold the area of the stroke that is in them"
    );
}

/// The walk buttons' chevron is two strokes, and a stroke is the one thing a mark drawn a whole
/// column at a time cannot be: the pixels its arms cross are some of the bar and some of the ink,
/// and a painter that gives each of them the one or the other draws a staircase (see
/// `draw_chevron`).
#[test]
fn the_chevrons_arms_carry_a_share_of_the_ink() {
    let side = 24usize;
    // A white bar under a dark ink, so that a pixel which is neither is a share of both: at full
    // coverage the pixel is the ink exactly, and at none of it the bar.
    let bar = [255u8, 255, 255];
    let ink = [56u8, 58, 66];

    let surface = DibSurface::create(side as u32, side as u32).expect("a surface");
    let filled = unsafe { std::slice::from_raw_parts_mut(surface.bits(), 4 * side * side) };
    for pixel in filled.as_chunks_mut::<4>().0 {
        pixel[0] = bar[2];
        pixel[1] = bar[1];
        pixel[2] = bar[0];
        pixel[3] = 255;
    }

    paint_glyph(
        &surface,
        CaptionButtonBox {
            kind: CaptionButton::Next,
            rect: RECT {
                left: 0,
                top: 0,
                right: side as i32,
                bottom: side as i32,
            },
        },
        false,
        ink,
        1.0,
    );

    // The two colours a pixel of the mark can be, in the buffer's own order, and the byte that
    // tells them apart: a pixel that is neither is carrying a share of the arms.
    let painted = surface.pixels();
    let partial = painted
        .as_chunks::<4>()
        .0
        .iter()
        .filter(|pixel| pixel[0] != ink[2] && pixel[0] != bar[2])
        .count();

    assert!(
        partial > 10,
        "the walk buttons' chevron carries part of the ink in {partial} pixels only: the pixels \
         its arms cross hold the area of the stroke that is in them"
    );
}

/// A caption is narrowed in steps, and each is the width the one before it stopped being worth it at:
/// the name is cut, then the name is given up, then the buttons go one at a time, and the bare bar is
/// the last of it.
///
/// It is one test rather than several because they are one rule read at several widths, and the thing
/// worth pinning down is the *order*: a name given up before it was cut would be a title that vanished
/// at the size it most needed to say something, and a caption given up while it still carried a row of
/// buttons that work on it would be a command bar thrown away with work left in it.
#[test]
fn a_caption_gives_up_its_name_before_it_gives_up_itself() {
    // Buttons are 46 logical pixels at 100% and are never shrunk, so a caption carries as many of them
    // as its width divides into 46, counted from close at the right edge back to the left (see
    // `button_boxes`), and its room for a name is its width less those buttons and less the 24 of
    // padding at either end of the run.
    assert_eq!(
        title_room(600, 96, true),
        600 - 9 * 46 - 24,
        "a caption wide enough carries the walk as well"
    );
    assert_eq!(
        title_room(240, 96, true),
        240 - 5 * 46 - 24,
        "and a narrower one carries only the five buttons that fit"
    );

    // The steps, at a width each of them is still worth taking at and one it is not.
    assert!(
        title_fits(600, 96, true) && caption_fits(600, 96, true),
        "a wide caption says the name of what is pinned"
    );
    assert!(
        !title_fits(240, 96, true) && caption_fits(240, 96, true),
        "a narrow one gives the name up and keeps the strip: it still carries five buttons, which is not a name's worth of room"
    );
    assert!(
        !title_fits(210, 96, true) && caption_fits(210, 96, true),
        "and one narrower still keeps its caption: four buttons and no name is a command bar, not a strip to give up"
    );
    assert!(
        caption_fits(20, 96, true),
        "and a caption too narrow for even close is a bare bar rather than no bar at all"
    );

    // A pin with nothing to maximize carries one window button fewer, so it has a button's width more
    // to say a name in: at 460 the caption with a maximize has given its name up and the one without
    // has not.
    assert!(
        !title_fits(460, 96, true) && title_fits(460, 96, false),
        "a caption's room is measured against the buttons it is actually carrying"
    );

    // And the whole rule is the display's own scale, because every width it is asked about is: a
    // caption on a display at 200% needs twice the room for the same name, and the window it is on is
    // measured in the same doubled pixels.
    assert!(!title_fits(240, 96, true) && !title_fits(480, 192, true));
    assert!(caption_fits(240, 96, true) && caption_fits(480, 192, true));
}

/// The strip a page's window carries says the page is read only, and the strips of every other kind
/// say nothing of it: the pill stands at the right end of the title area, immediately left of the
/// buttons (see `Caption::read_only` and `page_marks`).
///
/// Asked at 100% and at 150%, because every measurement the pill is drawn from is the display's scale
/// and a pill that only says its piece at 100% is one a reader at 150% never sees.
#[test]
fn a_page_says_it_is_read_only_and_a_caption_for_another_kind_does_not() {
    let palette = ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    };

    for dpi in [96u32, 144] {
        let scale = dpi as f32 / 96.0;
        let width = (600.0 * scale) as u32;
        let height = (30.0 * scale) as u32;
        let surface = DibSurface::create(width, height).expect("a surface");
        let strip = |read_only: bool| {
            let caption = Caption {
                title: "notes.txt",
                read_only,
                maximized: false,
                maximizable: true,
                always_on_top: true,
                hovered: None,
                pressed: None,
            };
            paint_caption(&surface, &palette, &caption, dpi);
            surface.pixels()
        };
        let page = strip(true);
        let picture = strip(false);

        let buttons = button_boxes(width as i32, height as i32, dpi, true);
        let title_right = buttons
            .iter()
            .map(|button| button.rect.left)
            .min()
            .unwrap_or(width as i32);
        let pill =
            page_marks::read_only_pill_box(&surface, width as i32, height as i32, title_right, dpi);
        assert!(
            pill.right - pill.left > 0,
            "a caption {width} wide has room for the pill at {dpi} dpi"
        );

        let background = palette.background;
        let ink_at = |pixels: &[u8], x: i32, y: i32| {
            let at = (y as usize * width as usize + x as usize) * 4;
            pixels[at..at + 3] != background
        };
        let said = |pixels: &[u8]| {
            (pill.left..pill.right)
                .flat_map(|x| (pill.top..pill.bottom).map(move |y| (x, y)))
                .any(|(x, y)| ink_at(pixels, x, y))
        };

        assert!(
            said(&page),
            "a page's caption says it is read only in the pill's own box, at {dpi} dpi"
        );
        assert!(
            !said(&picture),
            "and a caption that is not a page's says nothing there, at {dpi} dpi"
        );
    }
}

/// The name's room stops at the pill rather than running under it, and the room a name is measured
/// against is the room it actually has — a caption with no pill keeps a name the pill's room would
/// have given up (see `title_fits` and `title_fits_beside_pill`).
#[test]
fn a_name_stops_at_the_pill_and_is_only_given_up_where_the_pill_takes_its_room() {
    let palette = ChromePalette {
        background: [250, 250, 250],
        foreground: [30, 30, 30],
        accent: [10, 90, 200],
        dark: false,
    };

    let strip = |width: u32, title: &str, read_only: bool| {
        let surface = DibSurface::create(width, 30).expect("a surface");
        let caption = Caption {
            title,
            read_only,
            maximized: false,
            maximizable: true,
            always_on_top: true,
            hovered: None,
            pressed: None,
        };
        paint_caption(&surface, &palette, &caption, 96);
        surface.pixels()
    };

    // A caption wide enough for a name and the pill together, and one long enough to be cut.
    let width = 600u32;
    let surface = DibSurface::create(width, 30).expect("a surface");
    let buttons = button_boxes(width as i32, 30, 96, true);
    let title_right = buttons
        .iter()
        .map(|button| button.rect.left)
        .min()
        .unwrap_or(width as i32);
    let pill = page_marks::read_only_pill_box(&surface, width as i32, 30, title_right, 96);

    // The pill on its own, with nothing written beside it, so that what the name changed is the only
    // difference between the two strips.
    let pill_only = strip(width, "", true);
    let named = strip(
        width,
        "a rather long file name that has to be cut short.txt",
        true,
    );
    let background = palette.background;
    for y in 0..30i32 {
        for x in 0..width as i32 {
            let at = (y as usize * width as usize + x as usize) * 4;
            if named[at..at + 3] != pill_only[at..at + 3] {
                assert!(
                    x < pill.left,
                    "the name reached into the pill's own room at {x},{y}: the name's room is the \
                     pill's left edge, not the buttons'"
                );
            }
        }
    }

    // And a caption too narrow for both keeps the name: the pill is what the strip is saying about
    // the file, and a strip with no room to say what the file is has nothing for the pill to be
    // about. The width is found rather than written down, so the rule is asked rather than guessed.
    let narrow = (160..width as i32)
        .find(|width| title_fits(*width, 96, true) && !title_fits_beside_pill(*width, 96, true))
        .expect("a caption width where the pill's own room is the difference");
    let narrow = narrow as u32;
    let without_pill = strip(narrow, "notes.txt", false);
    let said = (0..narrow as i32).any(|x| {
        (0..30i32).any(|y| {
            let at = (y as usize * narrow as usize + x as usize) * 4;
            without_pill[at..at + 3] != background
        })
    });
    assert!(
        said,
        "a caption with no pill and no room for one still says the file's name: it is not the \
         pill's room it is giving up"
    );
}

/// The pill's ring closes at both ends: a run of text drawn on the caption's own surface is drawn with
/// an opaque background over the whole of its box, so a run whose box reached the pill's edge painted
/// that background over the ring's right cap and left the pill cut off where the words stop (see
/// `page_marks::paint_read_only_pill`).
///
/// Measured as ink rather than as a picture: the last pixel of the middle row is the ring, which is the
/// accent over the wash, so it is nearer the accent than the wash is. A run's background there would be
/// the wash itself, and the pill would read as open at that end.
#[test]
fn the_pills_ring_closes_at_the_end_the_words_stop_at() {
    let palette = ChromePalette {
        background: [30, 32, 40],
        foreground: [220, 220, 220],
        accent: [80, 170, 240],
        dark: true,
    };

    for dpi in [96u32, 144] {
        let scale = dpi as f32 / 96.0;
        let (width, height) = ((600.0 * scale) as u32, (30.0 * scale) as u32);
        let surface = DibSurface::create(width, height).expect("a surface");
        let buttons = button_boxes(width as i32, height as i32, dpi, true);
        let title_right = buttons
            .iter()
            .map(|button| button.rect.left)
            .min()
            .unwrap_or(width as i32);
        let box_ =
            page_marks::read_only_pill_box(&surface, width as i32, height as i32, title_right, dpi);
        page_marks::paint_read_only_pill(&surface, &palette, box_, dpi);

        let middle = (box_.top + box_.bottom) / 2;
        // The surface is the desktop's own byte order, so a pixel is read back the way it was
        // written: blue first.
        let at = |x: i32, y: i32| {
            let index = (y as usize * width as usize + x as usize) * 4;
            let pixels = surface.pixels();
            [pixels[index + 2], pixels[index + 1], pixels[index]]
        };
        let wash = text_paint::blend(palette.background, palette.accent, 0.14);
        let distance = |colour: [u8; 3], from: [u8; 3]| {
            let mut total = 0i32;
            for channel in 0..3 {
                let gap = colour[channel] as i32 - from[channel] as i32;
                total += gap * gap;
            }
            total
        };

        let edge = at(box_.right - 1, middle);
        assert!(
            distance(edge, palette.accent) < distance(edge, wash),
            "the pill's rightmost pixel at {dpi} dpi is the ring rather than the wash a run's \
             background leaves: it was {edge:?}"
        );
    }
}
