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

mod bubble;
mod caption;
mod menu;
mod transport;
