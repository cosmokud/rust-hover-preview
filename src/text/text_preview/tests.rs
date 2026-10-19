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

/// One painted line of a test's own: every character ten pixels wide, so a point can be put on a
/// character's cell by counting them.
fn a_line(text: &str) -> Vec<FrameLine> {
    vec![FrameLine {
        top: 0,
        height: 20,
        runs: vec![FrameRun {
            x: 0,
            width: text.chars().count() as i32 * 10,
            text: text.to_string(),
        }],
    }]
}

/// A double click takes the run of one kind of character around the one under the hand — letters,
/// digits and the underscore are a word, a gap is a run of space, and every other character is a
/// run of its own kind — and it is the character under the point that is taken, not the caret
/// beside it (see `word_in`).
#[test]
fn a_double_click_takes_the_run_of_the_character_under_the_hand() {
    let lines = a_line("let answer = 42; two_words -> done");
    let word = |character: usize| word_in(&lines, character as i32 * 10 + 5, 10).ordered();

    assert_eq!(
        word(9),
        ((0, 4), (0, 10)),
        "the right half of the last letter of a word still takes that word, not the gap after it"
    );
    assert_eq!(word(0), ((0, 0), (0, 3)), "a word is taken whole");
    assert_eq!(
        word(13),
        ((0, 13), (0, 15)),
        "a number is a word of its own kind"
    );
    assert_eq!(
        word(17),
        ((0, 17), (0, 26)),
        "an underscore is part of the word it is in"
    );
    assert_eq!(
        word(11),
        ((0, 11), (0, 12)),
        "a mark between two gaps is a run of its own"
    );
    assert_eq!(
        word(27),
        ((0, 27), (0, 29)),
        "and marks of one kind are taken together, as `->` is"
    );
    assert_eq!(word(3), ((0, 3), (0, 4)), "a gap is a run of its own");

    assert_eq!(
        word_in(&lines, 400, 10).ordered(),
        ((0, 30), (0, 34)),
        "a point past the end of a line takes the word at that end"
    );
    assert_eq!(
        word_in(&lines, -50, 10).ordered(),
        ((0, 0), (0, 3)),
        "and one before its start takes the word at that end"
    );

    assert_eq!(
        word_in(&a_line(""), 5, 10).ordered(),
        ((0, 0), (0, 0)),
        "a line with nothing on it has no word to take"
    );
    assert_eq!(
        word_in(&[], 5, 10).ordered(),
        ((0, 0), (0, 0)),
        "and neither has a page with no lines"
    );
}
