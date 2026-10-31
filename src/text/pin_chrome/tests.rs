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
mod menu;
mod transport;
