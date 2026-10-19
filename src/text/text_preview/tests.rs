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

/// The options a test's page is built with.
fn options() -> TextPreviewOptions {
    TextPreviewOptions {
        theme: TextTheme::Light,
        markdown_mode: MarkdownMode::Rendered,
        font_scale_percent: 125,
        full_mode: true,
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
        source: (0, 0),
    }]
}

/// A line too long for the box is drawn as several painted lines, and each of them says which
/// characters of the document line it is holding: a place in the preview is a place in the
/// document, so the tail of a wrapped line is the character the file has there rather than
/// whichever line of the frame it happens to be (see `FrameLine::source`).
#[test]
fn a_wrapped_line_says_which_characters_each_of_its_parts_holds() {
    let dir = scratch("wrapped");
    let source = dir.join("wrapped.txt");
    let long: String = (0..60).map(|n| format!("{n:02} ")).collect();
    fs::write(&source, format!("{long}\nsecond\n")).expect("a written page");

    let frame = render_scrolled(&source, 0, 140, 200, 96, options(), None).expect("a page");

    let wrapped: Vec<&FrameLine> = frame
        .lines
        .iter()
        .filter(|line| line.source.0 == 0)
        .collect();
    assert!(
        wrapped.len() > 1,
        "the line is wider than the box, so it is painted as several lines"
    );
    for pair in wrapped.windows(2) {
        assert!(
            pair[1].source.1 > pair[0].source.1,
            "every part of a wrapped line starts further into the document line than the part above it"
        );
    }

    // A point at the left edge of the second part is the character that part starts at, and not
    // the second line of the frame.
    let second = wrapped[1];
    let y = second.top + second.height / 2;
    assert_eq!(
        position_in(&frame.lines, 0, y),
        (0, second.source.1),
        "a press on a wrapped tail names the character the document has there"
    );
}

/// A copy is taken from the document rather than from the frame: a line too long for the box comes
/// back whole and unbroken, the two ends are cut to the characters the selection names, and a
/// selection reaching past the frame reaches past it here too (see `text_of`).
#[test]
fn a_copy_is_the_text_the_document_has() {
    let dir = scratch("copy");
    let source = dir.join("copy.txt");
    let long = "the quick brown fox jumps over the lazy dog and then it goes and does it again";
    fs::write(&source, format!("one\n{long}\ntwo\n")).expect("a written page");

    let all = Selection {
        anchor: (0, 0),
        caret: (usize::MAX, usize::MAX),
    };
    assert_eq!(
        text_of(&source, options(), all).as_deref(),
        Some(format!("one\r\n{long}\r\ntwo").as_str()),
        "the whole page is the file's own text, broken where the file breaks it and nowhere else"
    );

    let middle = Selection {
        anchor: (1, 4),
        caret: (1, 9),
    };
    assert_eq!(
        text_of(&source, options(), middle).as_deref(),
        Some("quick"),
        "a selection within a line is cut to the characters it names"
    );

    let across = Selection {
        anchor: (0, 1),
        caret: (2, 3),
    };
    assert_eq!(
        text_of(&source, options(), across).as_deref(),
        Some(format!("ne\r\n{long}\r\ntwo").as_str()),
        "a selection across lines joins them the way Windows does"
    );

    // The box this is copied out of does not hold the whole page: what a copy takes is the
    // document's, not the frame's.
    let frame = render_scrolled(&source, 0, 120, 40, 96, options(), None).expect("a page");
    assert!(
        frame.visible_lines < 3,
        "the box shows less of the page than the page has, so the copy above came from the document"
    );
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
