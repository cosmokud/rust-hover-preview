use super::*;

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
    assert_eq!(buttons.len(), 7);

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

/// A caption too narrow for the walk carries the window's buttons and nothing else. A
/// pin is worth less than a window it can be closed from, and half a walk is a set of
/// targets with no known order to them — so the whole group goes, and the buttons that
/// stay are exactly where they would have been on a caption wide enough to carry it.
#[test]
fn a_caption_narrower_than_its_own_walk_keeps_the_window_s_buttons() {
    let narrow = button_boxes(200, 30, 96, true);

    // The buttons are 46 wide either way here — 200 / 3 is 66, and the width is the
    // smaller of the two — so the walk's own four would need 7 * 46 = 322 of a strip
    // 200 wide, and the strip is the window's alone.
    assert_eq!(narrow.len(), 3);
    assert_eq!(narrow[0].kind, CaptionButton::Minimize);
    assert_eq!(narrow[1].kind, CaptionButton::Maximize);
    assert_eq!(narrow[2].kind, CaptionButton::Close);
    assert_eq!(narrow[2].rect.right, 200);

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

/// A pin with nothing to maximize — a sound's card — carries the two buttons that mean
/// something on it, packed against the right edge the way Windows packs a window's: the one
/// that closes it keeps the place it has on every other caption, and the one beside it is
/// the minimize that was there before. Neither of them moves for the walk, which sits
/// against the group and is measured off the same strip.
#[test]
fn a_caption_without_a_maximize_keeps_the_close_button_where_it_was() {
    let three = button_boxes(600, 30, 96, true);
    let two = button_boxes(600, 30, 96, false);

    assert_eq!(two.len(), 6);
    assert_eq!(two[0].kind, CaptionButton::Previous);
    assert_eq!(two[1].kind, CaptionButton::Next);
    assert_eq!(two[2].kind, CaptionButton::OpenWith);
    assert_eq!(two[3].kind, CaptionButton::OpenWithList);
    assert_eq!(two[4].kind, CaptionButton::Minimize);
    assert_eq!(two[5].kind, CaptionButton::Close);

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

    assert_eq!(two[5].rect.left, close.left);
    assert_eq!(two[5].rect.right, close.right);
    assert_eq!(two[4].rect.right, two[5].rect.left);
    assert_eq!(
        two[4].rect.right - two[4].rect.left,
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
            maximized: false,
            maximizable: true,
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
        maximized: false,
        maximizable: true,
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

/// The pictures this test writes are the design under review: the six marks the caption's
/// own buttons carry, at both scales, side by side in one strip. A row of marks is judged
/// as a row — one of them twice the size of the rest, or a third the size, or sitting a
/// pixel or two off the middle of its own button, is what a strip shows and a test on one
/// mark at a time cannot.
#[test]
fn draws_the_caption_glyphs() {
    let dir = scratch("caption-glyphs");
    // The marks on the bar's own background: the glyphs are ink, so a picture that is not
    // the colour behind them shows nothing, which is what a black buffer and black ink do.
    let bar = [250u8, 250, 250];
    let marks = [
        ("previous", CaptionButton::Previous, false),
        ("next", CaptionButton::Next, false),
        ("open-with", CaptionButton::OpenWith, false),
        ("open-with-list", CaptionButton::OpenWithList, false),
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
        let mut sizes: Vec<(&str, i32, i32)> = Vec::new();

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
                [56u8, 58, 66],
                scale,
            );

            let painted = surface.pixels();
            let ink = |x: usize, y: usize| painted[(y * side as usize + x) * 4] < 128;

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
            sizes.push((*name, tall, wide));

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
        for (name, tall, wide) in &sizes {
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
            .map(|(_, tall, _)| *tall);
        let tallest = heights.clone().max().unwrap_or_default();
        let smallest = heights.min().unwrap_or_default();
        assert!(
            tallest - smallest <= 2,
            "{tag} the row runs from {smallest} to {tallest} rows tall"
        );

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
#[test]
fn the_two_walk_buttons_are_chevrons_pointing_opposite_ways() {
    // A box with room for the whole glyph and a row to spare: the arms part half a span
    // either side of the middle row, so a box the glyph's own width would cut the open
    // end off and leave a test that passes for a chevron that is really a stub.
    const W: i32 = 24;
    const H: i32 = 24;
    let span = 10;
    let ink = [0u8, 0, 0];

    // The rows drawn in one column. The buffer is filled with a colour first, so that a
    // blank column reads as blank rather than as ink.
    fn drawn(buffer: &[u8], x: i32) -> Vec<i32> {
        (0..H)
            .filter(|y| buffer[((*y * W + x) * 4) as usize] == 0)
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
