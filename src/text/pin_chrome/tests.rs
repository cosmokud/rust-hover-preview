use super::*;

fn scratch(label: &str) -> std::path::PathBuf {
    let root = std::env::var_os("COMMANDCODE_SCRATCHPAD")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("pin-chrome")
        .join(label);
    std::fs::create_dir_all(&root).expect("a fixture directory");
    root
}

fn write_png(path: std::path::PathBuf, bgra: &[u8], width: u32, height: u32) {
    let mut rgba = bgra.to_vec();
    for pixel in rgba.as_chunks_mut::<4>().0 {
        pixel.swap(0, 2);
    }
    image::save_buffer(path, &rgba, width, height, image::ExtendedColorType::Rgba8)
        .expect("a written picture");
}

fn image_buffer(width: u32, height: u32) -> Vec<u8> {
    vec![0u8; width as usize * height as usize * 4]
}

/// One painted button's own width pasted into a strip of several, so that a row of glyphs is
/// laid out side by side the way the bar lays them out. The source is a button wide and the
/// destination is the whole strip, so the two are not the same stride.
fn blit_cell(out: &mut [u8], width: u32, height: u32, source: &[u8], source_width: u32, x: u32) {
    for row in 0..height as usize {
        for column in 0..source_width as usize {
            let from = (row * source_width as usize + column) * 4;
            let to = (row * width as usize + x as usize + column) * 4;
            if to + 3 < out.len() && from + 3 < source.len() {
                out[to..to + 4].copy_from_slice(&source[from..from + 4]);
            }
        }
    }
}

/// One painted strip drawn again at `factor` times the size, each pixel of it repeated that many
/// times, so that an edge half a pixel wide is a thing the eye can find: the marks this chrome
/// draws are ten pixels across, which is a size a picture of them has to be blown up to be read
/// at (see `draws_the_caption_glyphs`, `draws_the_marks_of_a_bar`).
fn zoom(buffer: &[u8], width: u32, height: u32, factor: u32) -> (Vec<u8>, u32, u32) {
    let (wide, tall) = (width * factor, height * factor);
    let mut out = vec![0u8; (wide * tall * 4) as usize];
    for y in 0..tall {
        for x in 0..wide {
            let from = ((y / factor) as usize * width as usize + (x / factor) as usize) * 4;
            let to = (y as usize * wide as usize + x as usize) * 4;
            out[to..to + 4].copy_from_slice(&buffer[from..from + 4]);
        }
    }
    (out, wide, tall)
}

mod bubble;
mod caption;
mod find_bar;
mod menu;
mod playlist;
mod transport;

#[test]
fn dpi_scale_sanitizes_zero_and_non_finite_factors() {
    let zero = DpiScale::from_dpi(0);
    assert_eq!(zero.dpi(), 96);
    assert_eq!(zero.factor(), 1.0);

    let hi = DpiScale::from_dpi(144);
    assert_eq!(hi.dpi(), 144);
    assert!((hi.factor() - 1.5).abs() < 1e-6);

    for bad in [0.0, -1.0, f32::NAN, f32::INFINITY] {
        let sanitized = DpiScale::from_factor(bad);
        assert_eq!(sanitized.dpi(), 96);
        assert_eq!(sanitized.factor(), 1.0);
    }

    // Zero DPI in caption/menu/transport/playlist never collapses to zero-size controls.
    let buttons = button_boxes(400, 30, 0, true);
    assert!(!buttons.is_empty());
    assert!(playlist_row_height(0.0) > 0);
}

#[test]
fn read_only_pill_and_chrome_layout_scale_with_fake_text_measurer() {
    let fake = primitives::FixedWidthMeasurer::new(8.0);
    let dpi_100 = DpiScale::from_dpi(96);
    let dpi_150 = DpiScale::from_dpi(144);
    let dpi_200 = DpiScale::from_dpi(192);

    let w100 = page_marks::read_only_pill_width_with(&fake, dpi_100);
    let w150 = page_marks::read_only_pill_width_with(&fake, dpi_150);
    let w200 = page_marks::read_only_pill_width_with(&fake, dpi_200);
    assert!(w100 > 0 && w150 > w100 && w200 > w150);

    // Wider font (12px/char vs 8px/char) must reduce title room beside the Read-Only pill.
    let wide_fake = primitives::FixedWidthMeasurer::new(12.0);
    let room_narrow = super::caption::title_room_beside_pill_with(600, dpi_150, true, &fake);
    let room_wide = super::caption::title_room_beside_pill_with(600, dpi_150, true, &wide_fake);
    assert!(room_wide < room_narrow);

    // Menu panel width and playlist name_travel both use TextMeasurer deterministically.
    let rows = vec![
        MenuRow {
            label: "Short".into(),
            mark: MenuMark::None,
            disabled: false,
        },
        MenuRow {
            label: "A much longer menu item label".into(),
            mark: MenuMark::None,
            disabled: false,
        },
    ];
    let menu_100 = super::menu::menu_main_panel_width_with(&rows, dpi_100, 1000, &fake);
    let menu_200 = super::menu::menu_main_panel_width_with(&rows, dpi_200, 1000, &fake);
    assert!(menu_200 > menu_100);

    assert!(super::playlist::name_travel_with("0123456789", 50, dpi_100, &fake).is_some());
    assert!(super::playlist::name_travel_with("0123", 50, dpi_100, &fake).is_none());
}
