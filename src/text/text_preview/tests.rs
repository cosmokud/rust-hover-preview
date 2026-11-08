use super::frame::TextFrame;
use super::*;
use crate::config::config::{MarkdownMode, TextTheme};
use crate::text::text_paint::luminance;
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
        let frame =
            render_scrolled(path, 0, width, height, 96, options, None, None).expect("a page");

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
        left: 0,
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

    let frame = render_scrolled(&source, 0, 140, 200, 96, options(), None, None).expect("a page");

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

/// **A line's own indentation is drawn, and a find's mark starts on the character it names.**
///
/// The whitespace a document line begins with is part of the line: a find, a caret and a press are
/// all measured in the characters the file holds. A row that began at its first non-space character
/// put every mark that followed the indent as many characters too far along — a query for `janitor`
/// in `  janitor a human` marked `nitor a` instead (see `find_all` and `wrap_line`).
#[test]
fn a_line_keeps_its_indent_and_a_find_lands_on_the_word_it_names() {
    let dir = scratch("indented-find");
    let source = dir.join("indented.txt");
    fs::write(&source, "  janitor a human\n").expect("a written page");

    let options = options();
    let plain = render_scrolled(&source, 0, 320, 160, 96, options, None, None).expect("a page");
    let found = find_all(&source, options, "janitor", false, false);
    assert_eq!(
        found.len(),
        1,
        "the query is found once, past the indent the line begins with"
    );
    let marked = render_scrolled(
        &source,
        0,
        320,
        160,
        96,
        options,
        None,
        Some(FindHighlights {
            matches: found,
            current: Some(0),
        }),
    )
    .expect("a page with a find on it");

    let line = plain.lines.first().expect("a painted line");
    assert_eq!(
        line.runs.first().map(|run| run.text.as_str()),
        Some("  "),
        "the indentation is drawn rather than dropped"
    );
    let word = line
        .runs
        .iter()
        .find(|run| run.text.starts_with("janitor"))
        .expect("the word's own run");

    let columns: Vec<i32> = (0..plain.width as i32)
        .filter(|x| {
            (line.top..line.top + line.height).any(|y| {
                frame_pixel(&plain, *x as u32, y as u32)
                    != frame_pixel(&marked, *x as u32, y as u32)
            })
        })
        .collect();
    let first = columns.first().copied().expect("the match is marked");
    let last = columns.last().copied().expect("the match is marked");

    assert_eq!(
        first, word.x,
        "the mark begins on the character the search found, not inside the word"
    );
    assert!(
        last < word.x + word.width,
        "and it ends where the word does rather than running into the text after it"
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

    // A line selection is the line and the break at its end, so pasting it puts a line of its own
    // back where it was taken from.
    let line = Selection {
        anchor: (0, 0),
        caret: (1, 0),
    };
    assert_eq!(
        text_of(&source, options(), line).as_deref(),
        Some("one\r\n"),
        "a line taken whole is the line and the break that ends it"
    );

    // The line after the one a selection took is where it ends, and the last line of a file whose
    // text stops without a break has no line after it: what a press on it takes is the text the
    // line does have (see `line_in`).
    let unbroken = dir.join("unbroken.txt");
    fs::write(&unbroken, "one\ntwo").expect("a written page");
    let last = Selection {
        anchor: (1, 0),
        caret: (2, 0),
    };
    assert_eq!(
        text_of(&unbroken, options(), last).as_deref(),
        Some("two"),
        "the last line of a file with no break at its end is taken to its own end"
    );

    // The box this is copied out of does not hold the whole page: what a copy takes is the
    // document's, not the frame's.
    let frame = render_scrolled(&source, 0, 120, 40, 96, options(), None, None).expect("a page");
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

/// A third press takes the whole of the line the word under the hand is on, and it is the
/// document's line rather than the painted one the point landed on: a line the box broke into
/// several painted pieces is taken whole, all of them, and a point between two lines takes the
/// nearer one (see `line_in`).
#[test]
fn a_third_press_takes_the_whole_of_the_line_under_the_hand() {
    let painted = |top: i32, text: &str, source: (usize, usize)| FrameLine {
        top,
        height: 20,
        left: 0,
        runs: vec![FrameRun {
            x: 0,
            width: text.chars().count() as i32 * 10,
            text: text.to_string(),
        }],
        source,
    };

    assert_eq!(
        line_in(&a_line("hello world"), 10).ordered(),
        ((0, 0), (1, 0)),
        "the line is taken whole, from its first character to the break at its end"
    );

    // A long line the box drew as two painted pieces of the one document line, which is what a
    // wrapped line is: a press on either piece takes the line, not the piece.
    let wrapped = vec![
        painted(0, "hello wor", (0, 0)),
        painted(20, "ld there", (0, 9)),
    ];
    assert_eq!(
        line_in(&wrapped, 30).ordered(),
        ((0, 0), (1, 0)),
        "a line too long for the box is taken whole, tail and all"
    );

    let page = vec![painted(0, "one", (0, 0)), painted(20, "two", (1, 0))];
    assert_eq!(
        line_in(&page, 25).ordered(),
        ((1, 0), (2, 0)),
        "a point between two lines takes the nearer one, which here is the later"
    );
    assert_eq!(
        line_in(&page, 15).ordered(),
        ((0, 0), (1, 0)),
        "and a point nearer the line before takes that one"
    );
    assert_eq!(
        line_in(&[], 10).ordered(),
        ((0, 0), (0, 0)),
        "a page with no lines has no line to take"
    );
}

/// A drag that follows a double press is grown by whole runs of characters rather than by single
/// ones: the hand takes the word it has reached whole, the gap it is over and nothing of the word
/// after it, and a hand that has not left the word the press took leaves that word as it was — all
/// of it, since a press on the last letter of a word still takes the word (see
/// `extend_word_selection`).
#[test]
fn a_drag_after_a_double_click_takes_the_runs_it_reaches_whole() {
    let lines = a_line("one two three");
    // `two`, the way a double press on it takes it.
    let two = word_in(&lines, 55, 10);
    let grown =
        |character: i32| extend_word_selection(&lines, two, character * 10 + 5, 10).ordered();

    assert_eq!(
        grown(5),
        ((0, 4), (0, 7)),
        "a hand still on the word the press took leaves it whole"
    );
    assert_eq!(
        grown(6),
        ((0, 4), (0, 7)),
        "and one on its last letter, where a drag by the character would have cut it down"
    );
    assert_eq!(
        grown(10),
        ((0, 4), (0, 13)),
        "a hand that has reached another word takes the whole of it"
    );
    assert_eq!(
        grown(1),
        ((0, 0), (0, 7)),
        "and one that has gone back before the word takes the word before it whole"
    );
    assert_eq!(
        grown(7),
        ((0, 4), (0, 8)),
        "the gap after the word is taken whole, and the word after it is not"
    );
    assert_eq!(
        extend_word_selection(&lines, two, 400, 10).ordered(),
        ((0, 4), (0, 13)),
        "a hand past the end of the line takes the run at that end"
    );
    assert_eq!(
        extend_word_selection(&lines, two, -50, 10).ordered(),
        ((0, 0), (0, 7)),
        "and one before its start takes the run at that end"
    );

    // A hand that has left the line altogether takes the run it finds on the line it has reached,
    // which is what wants the places of a selection to compare in reading order rather than line by
    // line.
    let mut page = a_line("two");
    let mut below = a_line("three");
    below[0].top = 20;
    below[0].source = (1, 0);
    page.extend(below);
    assert_eq!(
        extend_word_selection(&page, word_in(&page, 5, 10), 15, 30).ordered(),
        ((0, 0), (1, 5)),
        "a drag onto the line below takes the whole of the word it finds there"
    );

    // A line with nothing on it has no run to take, and a drag from one begins and ends at the
    // empty place the hand started from.
    let blank = a_line("");
    assert_eq!(
        extend_word_selection(&blank, word_in(&blank, 5, 10), 5, 10).ordered(),
        ((0, 0), (0, 0)),
        "a drag from a line with nothing on it has nothing to take"
    );
}

/// A find is read of the document rather than of the frame, and it searches from just after what is
/// selected to the end of the document and then from the beginning back to where it started: the
/// next occurrence is found wherever it is, the occurrence the selection already names is not the
/// answer, and a word that appears once has no next one at all (see `find_next`).
#[test]
fn a_find_walks_the_document_from_just_after_what_is_selected() {
    let dir = scratch("find");
    let source = dir.join("find.txt");
    fs::write(&source, "alpha beta\ngamma alpha\ndelta alpha\nepsilon\n").expect("a written page");

    let at = |line: usize, from: usize| Selection {
        anchor: (line, from),
        caret: (line, from + 5),
    };

    assert_eq!(
        find_next(&source, options(), at(0, 0)),
        Some(at(1, 6)),
        "the next `alpha` is the one two lines down"
    );
    assert_eq!(
        find_next(&source, options(), at(1, 6)),
        Some(at(2, 6)),
        "and the one after that is the next one again"
    );
    assert_eq!(
        find_next(&source, options(), at(2, 6)),
        Some(at(0, 0)),
        "and from the last of them the document wraps round to the first"
    );
    assert_eq!(
        find_next(
            &source,
            options(),
            Selection {
                anchor: (1, 6),
                caret: (1, 11)
            }
        ),
        Some(at(2, 6)),
        "the occurrence the selection names is where the search starts, not what it finds"
    );
    assert_eq!(
        find_next(
            &source,
            options(),
            Selection {
                anchor: (3, 0),
                caret: (3, 7)
            }
        ),
        None,
        "a word that appears once in the document has no next one, and no page to go to"
    );
    assert_eq!(
        find_next(
            &source,
            options(),
            Selection {
                anchor: (0, 0),
                caret: (0, 0)
            }
        ),
        None,
        "and neither has nothing at all selected, which is nothing to look for"
    );

    // The search is exact: `alpha` is looked for as it is written, and a run of letters the document
    // holds once is nowhere to be found at all.
    assert_eq!(
        find_next(
            &source,
            options(),
            Selection {
                anchor: (0, 0),
                caret: (0, 5)
            }
        ),
        Some(at(1, 6)),
        "`alpha` is looked for as it is written"
    );
    assert_eq!(
        find_next(
            &source,
            options(),
            Selection {
                anchor: (1, 1),
                caret: (1, 5)
            }
        ),
        None,
        "and `amma`, which the document writes once, has no next one"
    );
}

/// A search lists every place the query is found, in the order the document is read, and it is the
/// query that is looked for rather than a selection: nothing found is an empty list rather than a
/// page that has moved, and an empty query is nothing to look for (see `find_all`).
#[test]
fn a_search_lists_every_place_the_query_is_found() {
    let dir = scratch("find-all");
    let source = dir.join("find-all.txt");
    fs::write(&source, "alpha beta\ngamma alpha\ndelta alpha\nepsilon\n").expect("a written page");

    let alpha = |line: usize, from: usize| Selection {
        anchor: (line, from),
        caret: (line, from + 5),
    };

    assert_eq!(
        find_all(&source, options(), "alpha", true, false),
        vec![alpha(0, 0), alpha(1, 6), alpha(2, 6)],
        "every occurrence, in the order the document is read"
    );
    assert_eq!(
        find_all(&source, options(), "omega", true, false),
        Vec::new(),
        "a query the document does not hold is nowhere to be found"
    );
    assert_eq!(
        find_all(&source, options(), "", true, false),
        Vec::new(),
        "and an empty query is nothing to look for"
    );
}

/// Whether a search minds case is one of the bar's two switches, and the document is searched as it
/// is written either way: the same query finds one occurrence with the switch on and all of them
/// with it off (see `find_all`).
#[test]
fn a_case_blind_search_matches_a_query_in_any_case() {
    let dir = scratch("find-all-case");
    let source = dir.join("case.txt");
    fs::write(&source, "Alpha alpha ALPHA\n").expect("a written page");

    assert_eq!(
        find_all(&source, options(), "alpha", true, false).len(),
        1,
        "a search that minds case finds the one occurrence written as asked"
    );
    assert_eq!(
        find_all(&source, options(), "alpha", false, false).len(),
        3,
        "and one that does not finds all three"
    );
    assert_eq!(
        find_all(&source, options(), "ALPHA", false, false).len(),
        3,
        "whichever case the query itself is written in"
    );
}

/// A whole-word search takes only the occurrences that stand as words: `cat` in `concat` and inside
/// a `cat_cat` run is not one, while the same query before a full stop is (see `find_all`).
#[test]
fn a_whole_word_search_takes_only_the_words() {
    let dir = scratch("find-all-word");
    let source = dir.join("words.txt");
    fs::write(&source, "cat concat cat_cat cat.\n").expect("a written page");

    assert_eq!(
        find_all(&source, options(), "cat", true, false).len(),
        5,
        "every occurrence of the letters, wherever they stand"
    );
    assert_eq!(
        find_all(&source, options(), "cat", true, true),
        vec![
            Selection {
                anchor: (0, 0),
                caret: (0, 3),
            },
            Selection {
                anchor: (0, 19),
                caret: (0, 22),
            },
        ],
        "the word alone and the one a full stop ends, and not the three inside longer runs"
    );
}

/// A query made of punctuation ends in no word character, so a whole-word search asks for no
/// boundary at that end and finds it wherever it stands — `->` is found inside `a->b` as well as on
/// its own (see `find_all` and `word_boundaries_hold`).
#[test]
fn a_whole_word_search_of_punctuation_asks_for_no_boundary_of_its_own() {
    let dir = scratch("find-all-punctuation");
    let source = dir.join("arrows.txt");
    fs::write(&source, "a->b -> c\n").expect("a written page");

    assert_eq!(
        find_all(&source, options(), "->", true, true),
        vec![
            Selection {
                anchor: (0, 1),
                caret: (0, 3),
            },
            Selection {
                anchor: (0, 5),
                caret: (0, 7),
            },
        ],
        "punctuation is not a word, so a query made of it looks for itself wherever it stands"
    );
}

/// The search advances past each occurrence it finds rather than a character at a time, so two
/// overlapping readings of the same run are one answer and not two (see `find_all`).
#[test]
fn a_search_advances_past_each_hit_rather_than_counting_it_twice() {
    let dir = scratch("find-all-overlap");
    let source = dir.join("runs.txt");
    fs::write(&source, "aaaa\n").expect("a written page");

    assert_eq!(
        find_all(&source, options(), "aa", true, false),
        vec![
            Selection {
                anchor: (0, 0),
                caret: (0, 2),
            },
            Selection {
                anchor: (0, 2),
                caret: (0, 4),
            },
        ],
        "`aaaa` holds two `aa`, not the three a character-by-character walk would find"
    );
}

/// A search reads the document from its first line to its last and then stops: the occurrence on the
/// last line is found, and the ones before it are not found a second time on the way round — a list
/// of every place is not a walk that wraps (see `find_all`).
#[test]
fn a_search_reads_the_whole_document_and_never_wraps_into_its_first_hit_again() {
    let dir = scratch("find-all-order");
    let source = dir.join("order.txt");
    fs::write(&source, "omega\nbeta\nomega\n").expect("a written page");

    assert_eq!(
        find_all(&source, options(), "omega", true, false),
        vec![
            Selection {
                anchor: (0, 0),
                caret: (0, 5),
            },
            Selection {
                anchor: (2, 0),
                caret: (2, 5),
            },
        ],
        "the last line is read, and the first is not read again after it"
    );
}

/// One pixel of a painted frame, in the frame's own bytes: four to a pixel, blue first.
fn frame_pixel(frame: &TextFrame, x: u32, y: u32) -> [u8; 4] {
    let at = ((y * frame.width + x) * 4) as usize;
    let pixel = &frame.pixels[at..at + 4];

    [pixel[0], pixel[1], pixel[2], pixel[3]]
}

/// A find's places are drawn into the frame rather than over it: the page without a find is not the
/// page with one, every place the query was found is marked, and the one the reader is on is drawn
/// apart from the rest — which is what lets the count be read off the page as well as off the card
/// (see `FindHighlights` and `paint`).
#[test]
fn a_find_highlights_every_place_and_the_current_one_apart() {
    let dir = scratch("find-highlight");
    let source = dir.join("highlight.txt");
    fs::write(&source, "needle here\nnothing\nneedle again\n").expect("a written page");

    let at = |line: usize, from: usize| Selection {
        anchor: (line, from),
        caret: (line, from + 6),
    };
    let matches = vec![at(0, 0), at(2, 2)];
    let render = |find: Option<FindHighlights>| {
        render_scrolled(&source, 0, 320, 160, 96, options(), None, find).expect("a page")
    };

    let plain = render(None);
    let first = render(Some(FindHighlights {
        matches: matches.clone(),
        current: Some(0),
    }));
    let second = render(Some(FindHighlights {
        matches,
        current: Some(1),
    }));
    let nothing = render(Some(FindHighlights::default()));

    let differing = |a: &TextFrame, b: &TextFrame| {
        a.pixels
            .iter()
            .zip(&b.pixels)
            .filter(|(a, b)| a != b)
            .count()
    };

    assert!(
        differing(&plain, &first) > 0,
        "the page with a find on it is not the page without one"
    );
    assert!(
        differing(&first, &second) > 0,
        "which match is current is drawn on the page"
    );
    assert_eq!(
        differing(&plain, &nothing),
        0,
        "and a find that found nothing is the page exactly as it was"
    );
}

/// **A find's mark is stronger than the selection band, and its glyphs survive it** — in both the
/// bundled themes. Neither theme names a selection or a find colour, so both are derived from the
/// page; a find that was only a slightly different wash of the same page read as part of the text
/// (see `paint` and `find_ink`).
#[test]
fn a_find_stands_out_from_the_selection_and_keeps_its_glyphs_readable() {
    let dir = scratch("find-strength");
    let source = dir.join("strength.txt");
    fs::write(&source, "needle here\nneedle again\n").expect("a written page");

    let at = |line: usize| Selection {
        anchor: (line, 0),
        caret: (line, 6),
    };

    // The band drawn across the first six characters of `line`: the most common pixel of the region
    // is the band, and the furthest pixel from it is the ink laid over it.
    let band = |frame: &TextFrame, line: usize| {
        let line = frame
            .lines
            .iter()
            .find(|painted| painted.source.0 == line)
            .expect("a painted line");
        let run = line.runs.first().expect("a run of text");
        let characters = run.text.chars().count() as i32;
        let right = run.x + 6 * run.width / characters.max(1);

        let mut counts: std::collections::HashMap<[u8; 4], usize> =
            std::collections::HashMap::new();
        for y in line.top..line.top + line.height {
            for x in run.x..right {
                *counts
                    .entry(frame_pixel(frame, x as u32, y as u32))
                    .or_default() += 1;
            }
        }
        let mut pixels: Vec<([u8; 4], usize)> = counts.into_iter().collect();
        pixels.sort_by_key(|(_, count)| std::cmp::Reverse(*count));
        let grey = |pixel: [u8; 4]| luminance([pixel[2], pixel[1], pixel[0]]);
        let background = grey(pixels[0].0);
        let ink = pixels
            .iter()
            .map(|(pixel, _)| (grey(*pixel) - background).abs())
            .max_by(f32::total_cmp)
            .expect("a pixel");

        (background, ink)
    };

    for theme in [TextTheme::Light, TextTheme::Dark] {
        let options = TextPreviewOptions { theme, ..options() };
        let render = |selection, find| {
            render_scrolled(&source, 0, 320, 160, 96, options, selection, find).expect("a page")
        };
        let searched = FindHighlights {
            matches: vec![at(0), at(1)],
            current: Some(1),
        };

        let selection = render(Some(at(0)), None);
        let marks = render(None, Some(searched));
        let plain = render(None, None);
        let page = band(&plain, 0).0;
        let (selection_band, _) = band(&selection, 0);
        let (found_band, found_ink) = band(&marks, 0);
        let (current_band, current_ink) = band(&marks, 1);

        assert!(
            (found_band - selection_band).abs() >= 0.10,
            "{theme:?}: a find's band is a real step away from the selection's, not the same wash"
        );
        assert!(
            (current_band - found_band).abs() >= 0.03,
            "{theme:?}: the current match is the colour itself, the rest are a wash of it"
        );
        assert!(
            (current_band - page).abs() > (found_band - page).abs(),
            "{theme:?}: the current match is drawn opaque rather than mixed towards the page"
        );
        for (kind, ink) in [("find", found_ink), ("current", current_ink)] {
            assert!(
                ink >= 0.20,
                "{theme:?}: the glyphs on a {kind} mark stand clear of the band they are drawn on"
            );
        }
    }
}

/// **A pinned page keeps a column down its left for its own line numbers, and a hover does not** —
/// whatever the file is: source, prose, or a document drawn from its Markdown. The pin is the page a
/// reader works in, where a line is referred to by its number; a hover is something looked at on the
/// way past (see `gutter_width` and the layout's gutter).
#[test]
fn a_pinned_page_keeps_a_column_for_its_line_numbers_and_a_hover_does_not() {
    let dir = scratch("gutter");
    let source = dir.join("numbered.rs");
    fs::write(
        &source,
        "fn main() {
    let answer = 42;
}
",
    )
    .expect("a source file");

    let prose = dir.join("prose.txt");
    fs::write(
        &prose,
        "a page of prose
with two lines
",
    )
    .expect("a page of prose");

    let document = dir.join("numbered.md");
    fs::write(
        &document,
        "# Heading

A paragraph, and a list:

- one
- two
",
    )
    .expect("a rendered document");

    let pinned = TextPreviewOptions {
        full_mode: true,
        ..options()
    };
    let hover = TextPreviewOptions {
        full_mode: false,
        ..options()
    };

    let text_column = |path: &Path, options: TextPreviewOptions| {
        render_scrolled(path, 0, 400, 300, 96, options, None, None)
            .expect("a page")
            .lines
            .first()
            .and_then(|line| line.runs.first().map(|run| run.x))
            .expect("a painted line of text")
    };

    for path in [&source, &prose, &document] {
        assert!(
            text_column(path, pinned) > text_column(path, hover),
            "a pinned page keeps a column for its numbers that the same file hovered does not: {}",
            path.display()
        );
    }
}

/// A number is drawn once per document line rather than once per painted line: a line too long for
/// the box is painted as several lines, and only the first of them carries the number, while the
/// line after it carries its own (see `paint` and `own_line`).
#[test]
fn a_wrapped_line_is_numbered_once() {
    let dir = scratch("gutter-wrap");
    let source = dir.join("wrapped.rs");
    let long = "let value = compute_something(with, arguments, that, run, past, the, edge);";
    fs::write(&source, format!("// a comment\n{long}\n\nfn tail() {{}}\n")).expect("a source file");

    let options = TextPreviewOptions {
        full_mode: true,
        ..options()
    };
    let frame = render_scrolled(&source, 0, 200, 400, 96, options, None, None).expect("a page");

    let column = frame
        .lines
        .first()
        .and_then(|line| line.runs.first().map(|run| run.x))
        .expect("a painted line of text");
    let background = frame_pixel(&frame, 1, 1);
    let ink_in_the_column = |line: &FrameLine| {
        let mut ink = 0;
        for y in line.top..line.top + line.height {
            for x in 0..column {
                if frame_pixel(&frame, x as u32, y as u32) != background {
                    ink += 1;
                }
            }
        }

        ink
    };

    let wrapped: Vec<&FrameLine> = frame
        .lines
        .iter()
        .filter(|line| line.source.0 == 1)
        .collect();
    assert!(
        wrapped.len() > 1,
        "the line is wider than the box, so it is painted as several lines"
    );
    assert!(
        ink_in_the_column(wrapped[0]) > 0,
        "the first of them carries the number of the line"
    );
    for tail in &wrapped[1..] {
        assert_eq!(
            ink_in_the_column(tail),
            0,
            "and a wrapped tail carries none, because the line has one number"
        );
    }

    let next = frame
        .lines
        .iter()
        .find(|line| line.source.0 == 2)
        .expect("the line after the wrapped one");
    assert!(
        ink_in_the_column(next) > 0,
        "the line after it carries its own number"
    );
}

/// A window is kept as long as the document is, and answered for a request longer than that.
///
/// The table a window is kept in can only answer a request a held window *covers*, and what a
/// window covers is what it holds: a request for more lines than the document has is styled at the
/// document's own length, and one kept at the length it was *asked* for covers nothing — every
/// later request styled the same lines again, which on a file whose lines take a while to style is
/// the whole of that work paid again per question (see `styled_window`). The window a hover asks
/// for is the rows the display has room for, which on a short file is more lines than there are,
/// so that is the request this asks twice.
#[test]
fn a_window_kept_for_a_short_document_answers_a_request_longer_than_it() {
    let dir = scratch("window-cache");
    let source = dir.join("short.txt");
    fs::write(&source, "one\ntwo\nthree\n").expect("a text file");

    let options = options();
    let document = super::document::document(&source, options).expect("a document");
    let theme = crate::text::text_theme::loaded(options.theme).expect("a theme");

    let asked =
        super::document::styled_window(&document.key, &document, theme, 0, 500).expect("a window");
    let asked_again =
        super::document::styled_window(&document.key, &document, theme, 0, 500).expect("a window");

    assert_eq!(
        asked.lines.len(),
        3,
        "a window holds no more lines than the document has"
    );
    assert!(
        std::sync::Arc::ptr_eq(&asked, &asked_again),
        "the same window asked for twice was styled twice: a window of {} lines does not cover a \
         request for 500",
        asked.lines.len()
    );
}

/// The page is a document, so its caret moves the way a document's does: a line at a time with the
/// column kept over the short lines in between, a character at a time across the ends of lines, a run
/// at a time by the page's own run rule, and the two pairs of ends — the line's and the document's
/// (see `Motion` and `moved`).
fn page_of(lines: &[&str]) -> Vec<String> {
    lines.iter().map(|line| line.to_string()).collect()
}

#[test]
fn a_vertical_move_keeps_the_column_over_a_short_line() {
    let page = page_of(&["abcdefgh", "ab", "abcdefgh"]);

    let (short, goal) = moved(&page, (0, 6), Motion::Down, None);
    assert_eq!(
        short,
        (1, 2),
        "a line too short for the column takes the caret to its own end"
    );
    assert_eq!(goal, Some(6), "and the column it was aiming at is kept");

    let (long, _) = moved(&page, short, Motion::Down, goal);
    assert_eq!(
        long,
        (2, 6),
        "the line below is long enough for the column again, which is the whole of what a goal \
         column is for: `↓` over a short line and `↓` again must not walk the caret to the margin"
    );
}

#[test]
fn a_character_move_crosses_the_end_of_a_line() {
    let page = page_of(&["one", "two"]);

    assert_eq!(
        moved(&page, (0, 3), Motion::Right, None),
        ((1, 0), None),
        "the character after a line's last is the first of the line below"
    );
    assert_eq!(
        moved(&page, (1, 0), Motion::Left, None),
        ((0, 3), None),
        "and the one before a line's first is the last of the line above"
    );
    assert_eq!(
        moved(&page, (0, 0), Motion::Left, None),
        ((0, 0), None),
        "the beginning of the document is as far back as the caret goes"
    );
    assert_eq!(
        moved(&page, (1, 3), Motion::Right, None),
        ((1, 3), None),
        "and its end as far on"
    );
}

#[test]
fn the_two_pairs_of_ends_are_the_line_s_and_the_document_s() {
    let page = page_of(&["one", "two", "three"]);

    assert_eq!(moved(&page, (1, 2), Motion::LineStart, None).0, (1, 0));
    assert_eq!(moved(&page, (1, 0), Motion::LineEnd, None).0, (1, 3));
    assert_eq!(
        moved(&page, (1, 2), Motion::DocumentStart, None).0,
        (0, 0),
        "the document's beginning is the line's begin where the line is the first"
    );
    assert_eq!(
        moved(&page, (1, 2), Motion::DocumentEnd, None).0,
        (2, 5),
        "and its end is the end of the last line, not the line past it"
    );
}

#[test]
fn a_screenful_keeps_the_column_over_the_lines_it_passes() {
    let page: Vec<String> = (0..30)
        .map(|line| format!("line {line:02} of thirty"))
        .collect();

    let (turned, goal) = moved(&page, (0, 6), Motion::ScreenDown(20), None);
    assert_eq!(turned, (20, 6), "a page turned lands twenty lines on");
    assert_eq!(goal, Some(6), "at the column it was standing in");

    assert_eq!(
        moved(&page, turned, Motion::ScreenUp(20), goal).0,
        (0, 6),
        "and a page back is where the page was turned from"
    );
}

#[test]
fn a_move_past_either_end_leaves_the_caret_where_it_is() {
    let page = page_of(&["one", "two"]);

    assert_eq!(
        moved(&page, (0, 1), Motion::Up, None),
        ((0, 1), Some(1)),
        "an up arrow on the first line has nowhere to go"
    );
    assert_eq!(
        moved(&page, (1, 1), Motion::Down, None),
        ((1, 1), Some(1)),
        "and a down arrow on the last"
    );
}

#[test]
fn a_move_that_is_not_vertical_puts_the_goal_column_down() {
    let page = page_of(&["abcdefgh", "ab", "abcdefgh"]);

    assert_eq!(
        moved(&page, (0, 6), Motion::Right, Some(6)),
        ((0, 7), None),
        "a character move leaves the caret where the reader put it, so the next vertical move \
         aims from there rather than from the column a previous run began at"
    );
}

/// A document's lines are kept, bounded and uncoloured: a page can only scroll `MAX_DOC_LINES` of a
/// file, so nothing past that is read, and a file nothing renders is sliced out of the text the parse
/// already holds rather than taken through the highlighter.
///
/// Both halves are the whole of what makes a caret's arrow cheap: without the bound a press on a long
/// log styled the file past the lines it could ever show, and without the plain road every press of
/// every arrow was a document through syntect (see `drawn_lines`).
#[test]
fn the_lines_a_caret_reads_are_bounded_and_are_the_files_own() {
    let reach = crate::text::text_preview::document::MAX_DOC_LINES;
    let root = scratch("drawn");
    let path = root.join("a-long-page.txt");
    let text: String = (0..(reach + 500))
        .map(|line| format!("line {line}\n"))
        .collect();
    fs::write(&path, &text).expect("a written page");

    let options = TextPreviewOptions {
        theme: TextTheme::Dark,
        markdown_mode: MarkdownMode::Source,
        font_scale_percent: 100,
        full_mode: true,
    };
    let lines = drawn_lines(&path, options);

    assert_eq!(
        lines.len(),
        reach,
        "a page reads the lines it can reach and no more: the rest is a document no frame can show"
    );
    assert_eq!(
        lines[0], "line 0",
        "and a file nothing renders is its own lines, taken whole"
    );
    assert_eq!(
        lines[reach - 1],
        format!("line {}", reach - 1),
        "to the last line it reaches"
    );
}

/// A page moved under the reader is the page the frame after it would have been painted as: the
/// frame already on screen with its rows shifted, and only the rows that came in from an edge
/// painted (see `PaintedFrame`).
///
/// Every pixel is compared, not just the lines: a band that is one row short leaves the page above it
/// a row of the page before, and a scrollbar's thumb is the one thing about a page that moves
/// *without* the text moving, so a shifted frame owes it painted too.
#[test]
fn a_page_scrolled_is_the_page_the_frame_after_it_would_have_painted() {
    let dir = scratch("shifted");
    let path = dir.join("scrolled.txt");
    let text: String = (0..400)
        .map(|line| format!("    let value_{line} = compute({line}) * 3;\n"))
        .collect();
    fs::write(&path, &text).expect("a page to scroll");

    let options = TextPreviewOptions {
        full_mode: true,
        ..options()
    };
    let (width, height) = (240, 200);

    fn painted(frame: &TextFrame) -> PaintedFrame<'_> {
        PaintedFrame {
            pixels: &frame.pixels,
            width: frame.width,
            height: frame.height,
            first_line: frame.first_line,
            lines: &frame.lines,
            selection: None,
            find: None,
        }
    }

    // Down the page, and back up it: a shift either way is the same rows moved towards the edge the
    // page is going.
    for (from, to) in [(0usize, 1usize), (0, 3), (3, 1)] {
        let before = render_scrolled(&path, from, width, height, 96, options, None, None)
            .expect("a page at the place it is scrolled from");

        let shifted_before = shifted_frame_count();
        let shifted = render_scrolled_held(
            Some(painted(&before)),
            &path,
            to,
            width,
            height,
            96,
            options,
            None,
            None,
        )
        .expect("the page moved");
        let fresh = render_scrolled(&path, to, width, height, 96, options, None, None)
            .expect("the page painted afresh");

        assert!(
            shifted_frame_count() > shifted_before,
            "the frame on screen was moved rather than painted again, {from} to {to}"
        );
        assert_eq!(
            shifted.lines, fresh.lines,
            "the rows land where the page after it puts them, {from} to {to}"
        );
        assert_eq!(
            shifted.scrollbar, fresh.scrollbar,
            "the thumb is drawn where the page after it draws it, {from} to {to}"
        );
        assert_eq!(
            shifted.pixels, fresh.pixels,
            "and every pixel is the pixel that page would have been painted, {from} to {to}"
        );
    }

    // A page painted with something else in it is not a page that can be moved: a selection is a wash
    // laid under the text, and the pixels of a page without one are not the pixels of a page with it.
    let plain = render_scrolled(&path, 0, width, height, 96, options, None, None).expect("a page");
    let selected = Selection {
        anchor: (0, 0),
        caret: (2, 8),
    };
    let shifted_before = shifted_frame_count();
    let shifted = render_scrolled_held(
        Some(painted(&plain)),
        &path,
        1,
        width,
        height,
        96,
        options,
        Some(selected),
        None,
    )
    .expect("a page with a selection on it");
    let fresh = render_scrolled(&path, 1, width, height, 96, options, Some(selected), None)
        .expect("the page painted afresh");

    assert_eq!(
        shifted_frame_count(),
        shifted_before,
        "a page whose selection has changed is painted rather than moved"
    );
    assert_eq!(
        shifted.pixels, fresh.pixels,
        "and it is the page it would have been painted as"
    );
}

/// A pinned page's caret stands where the glyph it precedes is painted, at every display scale, and
/// the room between the numbers and the text is the display's rather than a literal.
///
/// The gutter is the one part of a page whose arithmetic is pure geometry: the numbers are as wide as
/// the widest number the page can reach, the gap between them and the text is that gap *at the
/// display's own scale*, and the text — and so the caret that stands in it — begins after both. A
/// number written for 100% is a caret that is right at 96 dpi and a few pixels off the glyph at 144,
/// and nothing here fails at 96 (see `gutter_width` and `gutter_gap`).
#[test]
fn the_caret_of_a_pinned_page_stands_on_the_glyph_it_precedes_at_any_scale() {
    let dir = scratch("gutter-dpi");
    let source = dir.join("scaled.txt");
    let text: String = (0..40)
        .map(|line| format!("line {line} of a page drawn at every scale\n"))
        .collect();
    fs::write(&source, text).expect("a page to draw");

    let options = TextPreviewOptions {
        full_mode: true,
        ..options()
    };

    let room_between_the_numbers_and_the_text = |frame: &TextFrame| {
        let first = frame.lines.first().expect("a painted line");
        let text_column = first.runs.first().expect("a run of text").x;
        let background = frame_pixel(frame, frame.width - 1, frame.height - 1);
        let number_ink: Vec<i32> = (first.top..first.top + first.height)
            .flat_map(|y| (0..text_column).map(move |x| (x, y)))
            .filter(|(x, y)| frame_pixel(frame, *x as u32, *y as u32) != background)
            .map(|(x, _)| x)
            .collect();

        let rightmost = number_ink.iter().copied().max().expect("a number's ink");
        text_column - rightmost - 1
    };

    let mut gaps = Vec::new();
    for dpi in [96u32, 144, 192] {
        let frame =
            render_scrolled(&source, 0, 400, 300, dpi, options, None, None).expect("a page");
        let text_column = frame
            .lines
            .first()
            .and_then(|line| line.runs.first().map(|run| run.x))
            .expect("a painted line of text");

        // The caret stands before character 0 of the first line, which is the left edge of the run
        // the text starts in — the glyph's own column, not the frame's.
        let caret = caret_rect(&frame.lines, (0, 0)).expect("a caret in the page");
        assert_eq!(
            caret.left, text_column,
            "at {dpi} dpi the caret is where the text begins rather than where the numbers do"
        );

        gaps.push(room_between_the_numbers_and_the_text(&frame));
    }

    assert!(
        gaps[0] > 0 && gaps[1] > gaps[0] && gaps[2] > gaps[1],
        "the room between the numbers and the text is the display's own gap: {gaps:?}"
    );
}

/// A caret on an empty line stands where that line's text would begin, not at the frame's edge.
///
/// An empty line is a painted line with no run on it, so there is nothing on it for a place to be
/// measured against — and the place is the left edge of the column the text is drawn in, past the
/// numbers the page keeps beside it, rather than the frame's own edge, which is inside the numbers
/// (see `x_within_line` and `FrameLine::left`). It is the same answer at a display scale, where the
/// gutter is wider than the number standing in it (see `gutter_width`).
#[test]
fn a_caret_on_an_empty_line_stands_at_the_text_column_and_never_on_the_numbers() {
    let dir = scratch("gutter-empty-line");
    let source = dir.join("empty-lines.txt");
    fs::write(&source, "21 the line above\n\n\n22 the line below\n").expect("a page to draw");

    let options = TextPreviewOptions {
        full_mode: true,
        ..options()
    };

    for dpi in [96u32, 144] {
        let frame =
            render_scrolled(&source, 0, 400, 300, dpi, options, None, None).expect("a page");
        let background = frame_pixel(&frame, frame.width - 1, frame.height - 1);

        // The column the text is drawn in, said by a line that has some on it.
        let text_column = frame.lines[0].runs[0].x;
        assert!(
            text_column > 0,
            "the page keeps a column for its numbers before the text begins"
        );

        // Lines 1 and 2 of the file are the empty ones: painted, numbered, and with no run at all.
        let empty: Vec<&FrameLine> = frame
            .lines
            .iter()
            .filter(|line| line.source.0 == 1 || line.source.0 == 2)
            .collect();
        assert_eq!(
            empty.len(),
            2,
            "an empty line still takes a line of the page"
        );

        for line in empty {
            assert!(
                line.runs.is_empty(),
                "an empty line has nothing on it to measure a place against"
            );

            let caret = caret_rect(&frame.lines, (line.source.0, 0)).expect("a caret on the line");
            assert_eq!(
                caret.left, text_column,
                "at {dpi} dpi a caret on an empty line stands where the line's text would begin"
            );

            // The line's number is drawn in the gutter beside it, and the caret is past all of it:
            // no part of the bar is laid over the numbers.
            let number_ink_right = (line.top..line.top + line.height)
                .flat_map(|y| (0..text_column).map(move |x| (x, y)))
                .filter(|(x, y)| frame_pixel(&frame, *x as u32, *y as u32) != background)
                .map(|(x, _)| x)
                .max();
            assert!(
                number_ink_right.is_some_and(|number| caret.left > number),
                "at {dpi} dpi the caret stands clear of the number drawn beside it"
            );
        }
    }
}

/// A copy of a pinned page is the document's text and never the numbers drawn beside it.
///
/// The gutter is painted into the frame and nothing else: the lines a selection is cut from are the
/// document's own, so a reader who selects the whole page and copies it gets the file rather than the
/// file with a column of numbers in front of every line (see `text_of` and `drawn_lines`).
#[test]
fn a_copy_of_a_pinned_page_is_the_text_and_never_its_numbers() {
    let dir = scratch("gutter-copy");
    let source = dir.join("copied.txt");
    let text: String = (0..12)
        .map(|line| format!("line {line} of a page with numbers\n"))
        .collect();
    fs::write(&source, &text).expect("a page to copy");

    let options = TextPreviewOptions {
        full_mode: true,
        ..options()
    };
    let frame = render_scrolled(&source, 0, 400, 300, 96, options, None, None).expect("a page");
    let hovered = render_scrolled(
        &source,
        0,
        400,
        300,
        96,
        TextPreviewOptions {
            full_mode: false,
            ..options
        },
        None,
        None,
    )
    .expect("the same page hovered");
    assert!(
        frame.lines[0].runs[0].x > hovered.lines[0].runs[0].x,
        "the page keeps a column for its numbers, which is what makes the copy worth asking about"
    );

    let whole = Selection {
        anchor: (0, 0),
        caret: (usize::MAX, usize::MAX),
    };
    let copied = text_of(&source, options, whole).expect("the page's text");

    // A copy joins the lines the way Windows does, so the lines are what is compared.
    assert_eq!(
        copied.lines().collect::<Vec<&str>>(),
        text.lines().collect::<Vec<&str>>(),
        "a copy is the file as it was read, not the page as it was drawn: no number is in it"
    );
}

/// A find's marks are laid over the text and never over the numbers: the column the numbers stand in
/// is the same page before and after a search, whatever the query found.
#[test]
fn a_find_is_drawn_over_the_text_and_never_over_the_numbers() {
    let dir = scratch("gutter-find");
    let source = dir.join("searched.txt");
    fs::write(&source, "a first line\nneedle on this line\na third line\n")
        .expect("a page to search");

    let options = TextPreviewOptions {
        full_mode: true,
        ..options()
    };
    let plain = render_scrolled(&source, 0, 400, 300, 96, options, None, None).expect("a page");
    let found = render_scrolled(
        &source,
        0,
        400,
        300,
        96,
        options,
        None,
        Some(FindHighlights {
            matches: vec![Selection {
                anchor: (1, 0),
                caret: (1, 6),
            }],
            current: Some(0),
        }),
    )
    .expect("a page with its find on it");

    let column = found.lines[1]
        .runs
        .first()
        .map(|run| run.x)
        .expect("the line's text");
    assert!(
        column > 0,
        "the page keeps a column for its numbers before the text begins"
    );

    let mut marks_over_the_numbers = 0;
    let mut marks_over_the_text = 0;
    for y in found.lines[1].top..found.lines[1].top + found.lines[1].height {
        for x in 0..found.width {
            if frame_pixel(&plain, x, y as u32) == frame_pixel(&found, x, y as u32) {
                continue;
            }

            if (x as i32) < column {
                marks_over_the_numbers += 1;
            } else {
                marks_over_the_text += 1;
            }
        }
    }

    assert!(
        marks_over_the_text > 0,
        "the place the query was found is painted on the page"
    );
    assert_eq!(
        marks_over_the_numbers, 0,
        "and nothing of the mark reaches the numbers standing beside the text"
    );
}
