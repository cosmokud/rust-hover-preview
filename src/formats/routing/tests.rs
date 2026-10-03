use super::*;
use std::path::PathBuf;

/// Which kind each list's names belong to, by the section the list is written under.
///
/// The two keys of the text section are the only pair anywhere that shares one, and they are
/// the same kind: the extensions and the names are the two ways a file is a text file
/// (see [`named_as`], which asks both). Everything else is one section to one kind.
///
/// It is keyed by section rather than written into the table because which kind a list
/// belongs to is a fact about this app's kinds and not about a list: the `[ffmpeg]` list is
/// the same kind as the `[video]` one — which engine plays a name rather than what a name is
/// — and no row could say so without also having to say what a kind is.
const KIND_OF_SECTION: &[(&str, &str, PreviewType)] = &[
    ("archive", "extensions", PreviewType::Archives),
    ("audio", "extensions", PreviewType::Audio),
    ("calibre", "extensions", PreviewType::Calibre),
    ("design", "extensions", PreviewType::Design),
    ("ebook", "extensions", PreviewType::Ebook),
    ("ffmpeg", "extensions", PreviewType::Videos),
    ("font", "extensions", PreviewType::Fonts),
    ("image", "extensions", PreviewType::Images),
    ("libre", "extensions", PreviewType::Libre),
    ("magick", "extensions", PreviewType::Magick),
    ("office", "extensions", PreviewType::Document),
    ("peazip", "extensions", PreviewType::Peazip),
    ("text", "extensions", PreviewType::Text),
    ("text", "names", PreviewType::Text),
    ("vector", "extensions", PreviewType::Vector),
    ("video", "extensions", PreviewType::Videos),
];

/// Every list this app ships, as the paths that reach it: an extension is reached by a name
/// carrying it, and a text name is reached by the name itself, since that is the lookup the
/// list is written for (see `text_formats::lookup_name`).
///
/// It is walked from the table rather than written out list by list, which is what the table
/// is for: a list this test does not name is a list whose names nothing here says reach the
/// kind they are written for, which is the one drift in this app nothing else in the tree
/// would have caught (see `every_shipped_name_reaches_the_kind_its_list_is_written_for`).
fn shipped_lists(config: &AppConfig) -> Vec<(PreviewType, Vec<PathBuf>)> {
    let mut lists: Vec<(PreviewType, Vec<PathBuf>)> = Vec::new();

    for list in crate::formats::lists::LISTS {
        let kind = KIND_OF_SECTION
            .iter()
            .find(|(section, key, _)| *section == list.section && *key == list.key)
            .map(|(_, _, kind)| *kind)
            .unwrap_or_else(|| {
                panic!(
                    "`[{}] {}` is a list this test does not name, so nothing here says what \
                     kind its names belong to",
                    list.section, list.key
                )
            });

        // An extension is reached by a name carrying it, `preview.tar.gz` included — the
        // archive row's compound entry is reached the only way it can be, by a name that ends
        // in it. A name row has no extension at all and is reached by the name itself.
        let paths = match list.entries {
            crate::formats::lists::Entries::Name => {
                list.entries(config).iter().map(PathBuf::from).collect()
            }
            _ => list
                .entries(config)
                .iter()
                .map(|name| PathBuf::from(format!("preview.{name}")))
                .collect(),
        };

        match lists.iter_mut().find(|(held, _)| *held == kind) {
            Some((_, already)) => already.extend(paths),
            None => lists.push((kind, paths)),
        }
    }

    lists
}

/// The names two lists hold, and the kind the order gives each of them.
///
/// Each is a name that is deliberately in two lists, and the winner is the kind asked
/// first. Nothing else may be in two lists: a name that reaches two kinds is a file whose
/// preview depends on the order rather than on the name, which is the thing this table
/// exists to keep written down and small.
const SHARED: &[(&str, PreviewType)] = &[
    // A `.ts` and an `.mts` are a transport stream in the video list and TypeScript in the
    // text lists, and the content settles which: a file of either name that is not a
    // transport stream is the source the text lists claim.
    ("preview.ts", PreviewType::Text),
    ("preview.mts", PreviewType::Text),
    // A `.dif` is a DV stream in the video list and a Data Interchange Format spreadsheet
    // in the render engine's list. The video list is asked first, so the stream wins: a
    // spreadsheet of that name is previewed as the video it is not.
    ("preview.dif", PreviewType::Videos),
    // A `.cdr` is a drawing the render engine draws and a CorelDRAW container the design
    // list reads a thumbnail out of. The engine is asked first, because what a thumbnail
    // shows of a drawing is not a preview of it.
    ("preview.cdr", PreviewType::Libre),
    // A `.vhd` is a Virtual PC disk image the listing engine opens and VHDL source the
    // text lists read, and nothing in the app had ever written the pair down: the two
    // lists were asked in an order that settled it and no test said which. It is the
    // listing engine's, because that list is asked first — a name that is a language more
    // often than it is a disk image is the reading the order gets wrong, and moving it is
    // a line in this table rather than a reordering of the lists.
    ("preview.vhd", PreviewType::Peazip),
    // A `.mpc` is a Musepack sound and the persistent cache ImageMagick keeps for itself,
    // and the two are told apart from the front of the file: an image cache opens with the
    // `id=ImageMagick` the converter's own reader answers for (see the signature that names
    // `miff`), and a Musepack file with its own `MPCK` or `MP+`. A name-level answer is
    // what this table settles, and it is the sound's, because the sound list is asked
    // first — a file of neither shape is the one the order gets wrong, and that is a file
    // neither reader could draw.
    ("preview.mpc", PreviewType::Audio),
];

/// Every name this app ships reaches the kind its list is written for, and a name two
/// lists hold reaches the kind that table names.
///
/// It is the test that would have caught the two chains that had drifted apart: the hook
/// asked the image converter's list before the design list and the loader asked the design
/// list first, so a name in both was gated as one kind and drawn as another. Nothing this
/// app ships was in both — which is why it went unnoticed — and this is what says so.
#[test]
fn every_shipped_name_reaches_the_kind_its_list_is_written_for() {
    let config = AppConfig::default();
    let lists = shipped_lists(&config);
    let mut unexpected: Vec<String> = Vec::new();

    for (kind, paths) in &lists {
        for path in paths {
            let name = path
                .file_name()
                .and_then(|name| name.to_str())
                .expect("a file name");

            let owners = lists.iter().filter(|(_, held)| held.contains(path)).count();

            let expected = match SHARED.iter().find(|(shared, _)| *shared == name) {
                Some((_, winner)) => *winner,
                None if owners > 1 => {
                    unexpected.push(format!(
                        "{name}: {kind:?} list and another list both claim it, and it is not \
                         written down in SHARED"
                    ));
                    continue;
                }
                None => *kind,
            };

            let reached = kind_of(path, &config);
            if reached != Some(expected) {
                unexpected.push(format!(
                    "{name}: the {kind:?} list's, reached {reached:?}, expected {expected:?}"
                ));
                continue;
            }

            // And the kind it was reached as is one whose own row says so, which is the half
            // of this that used not to be testable at all: `named_as` is a second answer to
            // the same question — asked of one kind rather than of the order — and a kind
            // whose row was taken out of it, or pointed at the wrong row, would gate nothing
            // while the router still answered with it. The eleven `is_<kind>_file` functions
            // were that second answer, one per module, and nothing in the tree ever compared
            // the two; this compares them for every name this app ships.
            if let Some(reached) = reached {
                if !named_as(path, &config, reached) {
                    unexpected.push(format!(
                        "{name}: the router calls it {reached:?} and its own lists do not \
                         claim it, so nothing gated would be drawn"
                    ));
                }
            }
        }
    }

    assert!(
        unexpected.is_empty(),
        "names reached a kind their list is not written for:\n  {}",
        unexpected.join("\n  ")
    );
}

/// The half of the `Ebook` kind is the one answer two tables used to work out for
/// themselves, and the drift it can suffer is a file whose kind is the book's and whose
/// reader is the comic's — which nothing above either would notice, because the kind was
/// right and only the page underneath it was wrong.
///
/// It is asserted over every name the book list ships rather than over a sample, and both
/// ways round: each name reaches the half its own spelling settles, and the two spellings
/// that are a page and a comic of the same kind are told apart from each other. The first
/// alone would pass a table that answered every name with a comic.
#[test]
fn every_shipped_book_name_reaches_the_half_its_spelling_settles() {
    let config = AppConfig::default();

    // The names the shipped list holds that are a page and a comic, read off the list
    // rather than written here, so a name added to either list is covered by this without
    // being added to this.
    let expected: Vec<(String, Page)> = lists::EBOOK
        .entries(&config)
        .iter()
        .map(|name| {
            let page = matches!(name.trim_start_matches('.'), "pdf" | "pdfa" | "epdf");
            (name.clone(), if page { Page::Pdf } else { Page::Comic })
        })
        .collect();

    assert!(
        expected.len() >= 2,
        "the book list has to carry a page and a comic for this to be worth saying"
    );

    let mut wrong: Vec<String> = Vec::new();
    let mut saw_page = false;
    let mut saw_comic = false;

    for (name, expected) in &expected {
        let path = PathBuf::from(format!("book.{name}"));
        let reached = page_of(&path, &config, None);

        match reached {
            Page::Pdf => saw_page = true,
            Page::Comic => saw_comic = true,
        }

        if reached != *expected {
            wrong.push(format!(
                "`{name}` is a page's spelling, read as {reached:?}"
            ));
        }

        // And the kind the router reaches is the book's either way: the half is what
        // picks the reader, not whether the file is a book at all.
        assert_eq!(
            kind_of(&path, &config),
            Some(PreviewType::Ebook),
            "`{name}` is the book kind's whichever half it is"
        );
    }

    assert!(
        saw_page && saw_comic,
        "the list has to hold both halves or this proves nothing about telling them apart"
    );
    assert!(
        wrong.is_empty(),
        "a book's spelling read as the wrong half of the kind:\n  {}",
        wrong.join("\n  ")
    );
}

/// The half of a drawing is the second answer two tables used to work out for themselves,
/// and it is asked in both directions: the image list asks it to reach the drawing's kind
/// rather than the picture's, and the loader asks it to pick the reader. A name the two
/// answered differently about is a metafile handed to the browser engine, so the assertion
/// is that both directions agree for every shipped name — and that a name reaching either
/// kind from the *other* list is still the drawing it is.
#[test]
fn every_shipped_drawing_name_reaches_one_half_and_the_list_agrees() {
    let config = AppConfig::default();

    let mut wrong: Vec<String> = Vec::new();
    let mut saw_svg = false;
    let mut saw_replayed = false;

    // The drawings this app ships, read off the two lists that hold them rather than
    // written here, so a name added to either is covered by this without being added here.
    let shipped: Vec<(String, PreviewType)> = lists::VECTOR
        .entries(&config)
        .iter()
        .map(|name| (name.clone(), PreviewType::Vector))
        .chain(
            lists::IMAGE
                .entries(&config)
                .iter()
                .map(|name| (name.clone(), PreviewType::Images)),
        )
        .collect();

    for (name, listed_as) in shipped {
        let path = PathBuf::from(format!("drawing.{name}"));
        let half = drawing_of(&path);
        let reached = kind_of(&path, &config);

        let (expected_half, expected_kind) = match half {
            Drawing::Svg => {
                saw_svg = true;
                (Drawing::Svg, PreviewType::Vector)
            }
            Drawing::Replayed => {
                saw_replayed = true;
                (Drawing::Replayed, listed_as)
            }
        };

        assert_eq!(
            half, expected_half,
            "`{name}`'s own half, which is what both tables read"
        );

        if let Some(reached) = reached {
            if reached != expected_kind {
                wrong.push(format!(
                    "`{name}` is the {expected_kind:?} kind's half, reached {reached:?}"
                ));
            }
        }

        // A drawing is reached from the vector list, and from the image list only where a
        // hand-edited entry left it there — never the other way round, which is the whole
        // of what the image entry is for.
        if reached == Some(PreviewType::Vector) && listed_as == PreviewType::Images {
            wrong.push(format!(
                "`{name}` was reached as a drawing through the image list, which is the \
                 one direction that order allows"
            ));
        }
    }

    assert!(
        saw_svg && saw_replayed,
        "the shipped names have to carry both halves or this proves nothing about telling \
         them apart"
    );
    assert!(
        wrong.is_empty(),
        "a drawing's name reached a kind its half is not:\n  {}",
        wrong.join("\n  ")
    );
}

/// The walk narrows a chain and never reorders it: what a file can be asked of is what its
/// kind declares, less whatever this machine cannot answer with, in the order the kind
/// names them — so the reader asked first is the first of the chain that can answer.
#[test]
fn the_readers_a_file_is_asked_of_are_its_chain_narrowed() {
    for (name, kind) in [
        ("letter.docx", PreviewType::Document),
        ("help.chm", PreviewType::Peazip),
        ("book.epub", PreviewType::Calibre),
        ("drawing.cdr", PreviewType::Libre),
        ("shot.nef", PreviewType::Magick),
        ("photo.jpg", PreviewType::Images),
        ("film.mp4", PreviewType::Videos),
        ("notes.txt", PreviewType::Text),
    ] {
        let path = Path::new(name);
        let asked = readers_for(kind, path);
        let declared = chain(kind);

        assert!(
            asked.len() <= declared.len(),
            "`{name}` cannot be asked of more readers than {kind:?} declares"
        );

        // A subsequence rather than a set: every reader asked is one the chain names, and
        // they come in the order it names them.
        let mut rest = declared.iter();
        for reader in &asked {
            assert!(
                rest.any(|declared| declared == reader),
                "`{name}` is asked of {reader:?} out of the order {kind:?} declares"
            );
        }
    }

    assert_eq!(
        readers_for(PreviewType::Document, Path::new("letter.docx")).first(),
        chain(PreviewType::Document).first(),
        "the first reader of a chain that can answer is the one asked"
    );
}

/// The two questions this module answers are the same question but in one place, and that
/// place is the video list's two shared names: asked of a file that is not a transport
/// stream, the content is read and the text lists win; asked of a name the content has
/// already answered with, there is nothing to read and the video list's answer stands.
#[test]
fn a_name_the_content_answered_with_is_asked_of_the_lists_alone() {
    let config = AppConfig::default();
    let stream = Path::new("content.ts");

    assert_eq!(
        kind_of_name(stream, &config),
        Some(PreviewType::Videos),
        "a signature that named a transport stream has already settled what it is"
    );
    assert_eq!(
        kind_of(stream, &config),
        Some(PreviewType::Text),
        "and a file of that name which is not one is the source the text lists claim"
    );
}

/// The fold the pin's own buttons walk by is the one a user would draw: a picture a
/// converter develops is a picture, a book a converter converted and a page an engine
/// drew are both a document, and a drawing is a drawing whether it was reached through
/// the image list or the vector one.
#[test]
fn a_kind_folds_into_the_thing_a_user_would_call_it() {
    for (kind, expected) in [
        (PreviewType::Images, NavCategory::Images),
        (PreviewType::Magick, NavCategory::Images),
        (PreviewType::Videos, NavCategory::Video),
        (PreviewType::Audio, NavCategory::Audio),
        (PreviewType::Ebook, NavCategory::Documents),
        (PreviewType::Calibre, NavCategory::Documents),
        (PreviewType::Document, NavCategory::Documents),
        (PreviewType::Libre, NavCategory::Documents),
        (PreviewType::Archives, NavCategory::Archives),
        (PreviewType::Peazip, NavCategory::Archives),
        (PreviewType::Text, NavCategory::Text),
        (PreviewType::Fonts, NavCategory::Fonts),
        (PreviewType::Design, NavCategory::Design),
        (PreviewType::Vector, NavCategory::Design),
    ] {
        assert_eq!(nav_category(kind), expected, "{kind:?} is a {expected:?}");
    }
}

/// A kind and the kind it shares a switch with are one category, always: the two answers
/// `All` and `Category` give a folder are supposed to differ, and a pair of kinds the tray
/// cannot switch apart is a difference they cannot have.
#[test]
fn the_kinds_that_share_a_switch_share_a_category() {
    let config = AppConfig::default();
    let kinds = [
        PreviewType::Images,
        PreviewType::Magick,
        PreviewType::Videos,
        PreviewType::Audio,
        PreviewType::Ebook,
        PreviewType::Calibre,
        PreviewType::Document,
        PreviewType::Libre,
        PreviewType::Archives,
        PreviewType::Peazip,
        PreviewType::Text,
        PreviewType::Fonts,
        PreviewType::Design,
        PreviewType::Vector,
    ];

    // The whole set is switched off one kind at a time, and every kind whose own gate
    // went down with it is the kind that shares its switch — which is the fold's own
    // table, read back off the gates rather than off the arms that wrote it.
    for kind in kinds {
        let mut probe = config.clone();
        probe.image_preview_enabled = false;
        probe.video_preview_enabled = false;
        probe.audio_preview_enabled = false;
        probe.text_preview_enabled = false;
        probe.ebook_preview_enabled = false;
        probe.archive_preview_enabled = false;
        probe.document_preview_enabled = false;
        probe.font_preview_enabled = false;
        probe.design_preview_enabled = false;
        probe.vector_preview_enabled = false;
        kind.set_enabled_in(&mut probe, true);

        let switched = kind.enabled_in(&probe);
        assert!(
            switched,
            "{kind:?} is switched back on by the switch it is written for"
        );

        // What one switch turns on is one category: the tray can only narrow a walk by
        // something it can also switch, so a gate that brought up a kind of another
        // category would put a file in a walk the user had turned off.
        for other in kinds {
            if other.enabled_in(&probe) {
                assert_eq!(
                    nav_category(other),
                    nav_category(kind),
                    "{other:?} is switched by {kind:?}'s gate, and so has to be walked with it"
                );
            }
        }
    }
}

/// The chain names the readers of each kind, and the two names whose preview was given up
/// on purpose keep the single reader they have: a `.cdr` is a page an engine draws and
/// nothing without one, and a `.chm` is a listing that is there immediately rather than a
/// page that takes two seconds to draw.
#[test]
fn the_chain_names_the_readers_of_each_kind() {
    let config = AppConfig::default();

    assert_eq!(
        chain(PreviewType::Document),
        &[Reader::Office, Reader::LibreOffice],
        "a document is the application's to draw where it is installed, and the engine's where it is not"
    );
    assert_eq!(
        chain(PreviewType::Images),
        &[Reader::Native],
        "a picture has one reader, which is the file's own question"
    );

    for (name, kind) in [
        ("drawing.cdr", PreviewType::Libre),
        ("help.chm", PreviewType::Peazip),
    ] {
        let reached = kind_of(Path::new(name), &config);
        assert_eq!(reached, Some(kind), "`{name}` is the {kind:?} kind's");
        assert_eq!(
            chain(kind).len(),
            1,
            "`{name}` has the one reader it has always had: what it gives up is deliberate"
        );
    }
}

/// A sound is reached by its name, and a container whose streams a probe found a sound in
/// and no picture is reached by that verdict — the one thing a list cannot say, and the
/// one thing that takes a file away from the video claim above it.
#[test]
fn a_sound_is_the_name_it_carries_or_the_streams_a_probe_found() {
    let config = AppConfig::default();

    for name in [
        "track.flac",
        "song.mp3",
        "book.m4b",
        "radio.mka",
        "album.opus",
    ] {
        assert_eq!(
            kind_of(Path::new(name), &config),
            Some(PreviewType::Audio),
            "`{name}` is a sound"
        );
    }

    assert_eq!(
        chain(PreviewType::Audio),
        &[Reader::Native, Reader::Ffmpeg],
        "a sound is played inside this app's own process before FFmpeg's player is asked"
    );

    // A container the video list claims, whose own streams the probe found no picture in:
    // the verdict is what answers it, and it is remembered under the file it was read from
    // rather than under its name — so the path here is one no other test writes.
    let probed = Path::new("routing-test-audio-only-container.mkv");
    assert_eq!(
        kind_of(probed, &config),
        Some(PreviewType::Videos),
        "a container of a video's name is a video until a probe says otherwise"
    );

    crate::formats::audio_formats::remember_audio_only(probed);
    assert_eq!(
        kind_of(probed, &config),
        Some(PreviewType::Audio),
        "and a sound once its own streams have answered for it"
    );
    assert_eq!(
        kind_of_name(probed, &config),
        Some(PreviewType::Audio),
        "the same answer through the name half of the question, which is the form the \
         content tier asks in"
    );
}

/// One route and the four questions it is made of are the same four answers, for every
/// name this app ships.
///
/// `resolve` exists so that a caller asks once and reads fields, and what that is worth is
/// only true if the record cannot say something the functions it replaced would have
/// disagreed with — so this asks both, over the shipped lists rather than over a sample, and
/// compares. The kind and the half of each kind are the two that were reached from different
/// modules, which is how a book came to be drawn by the reader the comic's or a drawing by
/// the browser engine's: the kind was right and the reader was another kind's (see
/// [`every_shipped_drawing_name_reaches_one_half_and_the_list_agrees`]).
///
/// Who draws it is asserted too, since that is the answer none of the four has on its own:
/// a kind this app ships is drawn by something for every name, and the four engine kinds
/// are the ones whose answer is `Nothing` — the engine's own window is not a reader of this
/// app's, which is what `chain` is for (see [`DrawnBy`]).
#[test]
fn a_route_says_what_the_four_questions_it_is_made_of_say() {
    let config = AppConfig::default();

    for (kind, paths) in shipped_lists(&config) {
        for path in paths {
            let probe = crate::formats::content_type::Probe::read(&path);
            let route = resolve(&path, &config, &probe);

            assert_eq!(
                route.named,
                kind_of(&path, &config),
                "`{}` is reached as one kind here and another by the router",
                path.display()
            );
            assert_eq!(
                route.drawing,
                drawing_of(&path),
                "`{}` is one half of the drawing kind here and another by the name",
                path.display()
            );

            if route.named == Some(PreviewType::Ebook) {
                assert_eq!(
                    route.page,
                    page_of(&path, &config, probe.facts()),
                    "`{}` is one half of the book kind here and another by the list",
                    path.display()
                );
            } else {
                assert_eq!(
                    route.page,
                    Page::Comic,
                    "`{}` is not a book, so which half of one it is has no answer to be \
                     wrong about",
                    path.display()
                );
            }

            // The four engine kinds and nothing else are drawn by an engine's own window
            // rather than by a reader of this app's, so they are the four arms that answer
            // `Nothing` — the only way a claimed kind has no hand to be drawn by.
            let engine_kind = matches!(
                route.named,
                Some(PreviewType::Document)
                    | Some(PreviewType::Libre)
                    | Some(PreviewType::Calibre)
                    | Some(PreviewType::Magick)
            );

            assert_eq!(
                route.drawn_by == DrawnBy::Nothing,
                engine_kind,
                "`{}` is the {kind:?} list's, and its reader is {}",
                path.display(),
                if engine_kind {
                    "an engine's own window rather than a reader of this app's"
                } else {
                    "a reader of this app's own"
                }
            );
        }
    }
}
