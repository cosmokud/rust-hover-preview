use super::frame::TextFrame;
use super::*;
use crate::config::config::{MarkdownMode, TextTheme};
use std::fs;
use std::path::{Path, PathBuf};

/// Where the fixtures and the pictures of them are written. The scratchpad
/// the session hands out, so a render can be looked at rather than only
/// asserted.
fn scratch(label: &str) -> PathBuf {
    let root = std::env::var_os("COMMANDCODE_SCRATCHPAD")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("text-preview")
        .join(label);
    fs::create_dir_all(&root).expect("a fixture directory");
    root
}

fn write_png(dir: &Path, name: &str, frame: &TextFrame) {
    let mut rgba = frame.pixels.clone();
    for pixel in rgba.as_chunks_mut::<4>().0 {
        pixel.swap(0, 2);
    }
    image::save_buffer(
        dir.join(format!("{name}.png")),
        &rgba,
        frame.width,
        frame.height,
        image::ExtendedColorType::Rgba8,
    )
    .expect("a written picture");
}

/// The painting a text preview does is shared with archive listings, so the
/// path is worth a smoke test of its own: a page is measured, painted, and
/// opaque, and what it paints is the frame the compositor gets.
#[test]
fn draws_pages_of_text() {
    let dir = scratch("pages");
    let source = dir.join("sample.rs");
    fs::write(
        &source,
        "fn main() {\n    // a comment\n    let answer = 42;\n    println!(\"{answer}\");\n}\n",
    )
    .expect("a source file");

    let markdown = dir.join("sample.md");
    fs::write(
        &markdown,
        "# Heading\n\nSome *emphasis*, a [link](https://example.com), and `code`.\n\n```rust\nlet x = 1;\n```\n",
    )
    .expect("a document");

    for (path, name, full_mode, theme) in [
        (&source, "source-light", false, TextTheme::Light),
        (&markdown, "markdown-dark", true, TextTheme::Dark),
    ] {
        let options = TextPreviewOptions {
            theme,
            markdown_mode: MarkdownMode::Rendered,
            font_scale_percent: 125,
            full_mode,
        };

        let (width, height) = measure(path, 1_920, 1_200, 96, options).expect("a measured page");
        let frame = render_scrolled(path, 0, width, height, 96, options, None).expect("a page");

        assert!(frame.width > 0 && frame.height > 0, "{name}");
        assert!(
            frame
                .pixels
                .as_chunks::<4>()
                .0
                .iter()
                .all(|pixel| pixel[3] == 255),
            "{name} is a page, so every pixel of it is opaque"
        );

        write_png(&dir, name, &frame);
    }
}
