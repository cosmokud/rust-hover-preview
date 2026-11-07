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
    ("f3d", "extensions", PreviewType::Models),
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
        let kind = kind_of_section(list);

        // An extension is reached by a name carrying it, `preview.tar.gz` included — the
        // archive row's compound entry is reached the only way it can be, by a name that ends
        // in it. A name row has no extension at all and is reached by the name itself, and by
        // the dotted one as well, since a `.env` and an `env` are one entry to the rule that
        // reads it.
        let paths = match list.entries {
            crate::formats::lists::Entries::Name => list
                .entries(config)
                .iter()
                .flat_map(|name| [PathBuf::from(name), PathBuf::from(format!(".{name}"))])
                .collect(),
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

/// The kind a row's names belong to, or a panic where this test has no answer for it: a list
/// added to the table without being named here is a list nothing says the kind of.
fn kind_of_section(list: &crate::formats::lists::List) -> PreviewType {
    KIND_OF_SECTION
        .iter()
        .find(|(section, key, _)| *section == list.section && *key == list.key)
        .map(|(_, _, kind)| *kind)
        .unwrap_or_else(|| {
            panic!(
                "`[{}] {}` is a list this test does not name, so nothing here says what \
                 kind its names belong to",
                list.section, list.key
            )
        })
}

/// The kinds whose own lists claim `path`, asked of each row the way the row itself reads a
/// name.
///
/// The rows are asked rather than their paths compared, because two lists can reach one file
/// under two names of their own: a `[text] names` entry and an extension list are two lookups
/// for the same file, and the archive row's `tar.gz` reaches a name whose extension the listing
/// list carries. Both are overlaps this table owes an answer for, and neither is a pair of equal
/// paths.
fn kinds_claiming(path: &Path, config: &AppConfig) -> Vec<PreviewType> {
    let mut kinds: Vec<PreviewType> = Vec::new();

    for list in crate::formats::lists::LISTS {
        if list.claims(path, config) {
            let kind = kind_of_section(list);

            if !kinds.contains(&kind) {
                kinds.push(kind);
            }
        }
    }

    kinds
}

/// The names two lists hold, and the kind the order gives each of them.
///
/// Each is a name that is deliberately in two lists, and the winner is the kind asked
/// first. Nothing else may be in two lists: a name that reaches two kinds is a file whose
/// preview depends on the order rather than on the name, which is the thing this table
/// exists to keep written down and small.
///
/// What is written down here is the *name's* answer. A file of one of these names may be
/// settled by its own bytes before the order is ever asked — a `.ts` source, a `.vhd` of VHDL,
/// a `.dif` that is a spreadsheet — and that half belongs to `content_type`, which every
/// caller that draws a file asks of it (see `routing::resolve` and `probe::GUARDS`).
const SHARED: &[(&str, PreviewType)] = &[
    // A `.ts` and an `.mts` are a transport stream in the video list and TypeScript in the
    // text lists, and the content settles which: a file of either name that is not a
    // transport stream is the source the text lists claim.
    ("preview.ts", PreviewType::Text),
    ("preview.mts", PreviewType::Text),
    // A `.dif` is a DV stream in the video list and a Data Interchange Format spreadsheet
    // in the render engine's list. The video list is asked first, so the name is the
    // stream's — and the spreadsheet is the formats table's rather than the order's: a
    // `.dif` whose head reads as the text its records are written in is the render
    // engine's before the order is asked at all (see `spreadsheet_or_dv_stream`).
    ("preview.dif", PreviewType::Videos),
    // A `.vhd` is a Virtual PC disk image the listing engine opens and VHDL source the
    // text lists read, and the listing engine's is the name's answer, because that list is
    // asked first. The language is the reading the order gets wrong and it is the common
    // one, so the file's own head is what settles it: a `.vhd` that reads as text is a page
    // of text, and one that is the disk image it is spelled like — a header of `conectix`, a
    // fixed image's raw data — keeps the listing (see `vhdl_source_or_disk_image`).
    ("preview.vhd", PreviewType::Peazip),
    // A `.mpc` is a Musepack sound and the persistent cache ImageMagick keeps for itself,
    // and the two are told apart from the front of the file: an image cache opens with the
    // `id=ImageMagick` the converter's own reader answers for (see the signature that names
    // `miff`), and a Musepack file with its own `MPCK` or `MP+`. A name-level answer is
    // what this table settles, and it is the sound's, because the sound list is asked
    // first — and a file whose bytes named the converter is the converter's whatever the
    // order says (see `HoverFacts::is_audio`).
    ("preview.mpc", PreviewType::Audio),
    // A `.tar.gz` is a whole file's name to the archive list, which reads its dotted entry
    // for what it is, and a `gz` stream to the listing list, which reads the extension the
    // name ends in. The archive list is asked first, which is what keeps a tarball out of
    // the engine's hands (see `peazip_formats::is_peazip_file`).
    ("preview.tar.gz", PreviewType::Archives),
    // A `.dxf` is a drawing the render engine imports and a scene the viewer reads the 3D half
    // of, and the two are told apart by what the file holds rather than by the name: the one
    // written down here is a sheet, which is what a file of the name that holds no mesh is
    // (`preview.dxf` is not there at all, so this is the answer both reach). See `claim_libre`
    // for the split and `model_formats::dxf` for the geometry that settles it.
    ("preview.dxf", PreviewType::Libre),
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
    let config = AppConfig {
        text_dont_preview_secrets: false,
        ..Default::default()
    };
    let lists = shipped_lists(&config);
    let mut unexpected: Vec<String> = Vec::new();

    for (kind, paths) in &lists {
        for path in paths {
            let name = path
                .file_name()
                .and_then(|name| name.to_str())
                .expect("a file name");

            let owners = kinds_claiming(path, &config);

            let expected = match SHARED.iter().find(|(shared, _)| *shared == name) {
                Some((_, winner)) => *winner,
                None if owners.len() > 1 => {
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

    // And every name written down in the table is still one two lists hold. An entry whose name
    // has left one of the two is an exception nothing would notice the table no longer needs —
    // it asserts the same kind the order already gives — and the table is meant to be as small
    // as the overlaps it records. It is how a `.cdr` entry outlived the list that carried the
    // second name.
    let stale: Vec<String> = SHARED
        .iter()
        .map(|(shared, _)| (*shared, kinds_claiming(Path::new(shared), &config)))
        .filter(|(_, owners)| owners.len() < 2)
        .map(|(shared, owners)| format!("{shared}: only {owners:?} claims it now"))
        .collect();

    assert!(
        stale.is_empty(),
        "names in SHARED that are no longer in two lists:\n  {}",
        stale.join("\n  ")
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

/// The picture list is the master bucket for everything Windows decodes that *is* a picture, and
/// only for those: the ten spellings of the codecs beside the ones that were always there, and the
/// camera raws, the cursor and the run-length bitmap the engine's picture converter was the only
/// home for. What is not a picture is another kind's however an engine draws it — a layered
/// document, a scanner's dataset, a page of rich text — and that is the half of the rule this
/// test is as much about as the names it does move.
#[test]
fn the_pictures_windows_decodes_are_the_picture_lists() {
    let config = AppConfig::default();

    for name in [
        // The ten spellings of the codecs there already were.
        "picture.dib",
        "picture.exif",
        "picture.heics",
        "picture.heifs",
        "picture.avcs",
        "picture.avifs",
        "picture.hif",
        "picture.icon",
        "picture.jxr",
        "picture.wdp",
        // And the camera raws, the cursor and the run-length bitmap.
        "shot.3fr",
        "shot.arw",
        "shot.cr2",
        "shot.cr3",
        "shot.crw",
        "shot.dcr",
        "shot.dng",
        "shot.erf",
        "shot.fff",
        "shot.iiq",
        "shot.k25",
        "shot.kdc",
        "shot.mef",
        "shot.mos",
        "shot.mrw",
        "shot.nef",
        "shot.nrw",
        "shot.orf",
        "shot.pef",
        "shot.raf",
        "shot.raw",
        "shot.rw2",
        "shot.rwl",
        "shot.sr2",
        "shot.srf",
        "shot.srw",
        "shot.x3f",
        "cursor.cur",
        "bitmap.rle",
    ] {
        let path = Path::new(name);

        assert_eq!(
            kind_of_name(path, &config),
            Some(PreviewType::Images),
            "`{name}` is a picture, and its name is what says so"
        );
        assert!(
            named_as(path, &config, PreviewType::Images),
            "`{name}` is the picture list's"
        );
        assert!(
            !named_as(path, &config, PreviewType::Magick),
            "and not the image converter's as well: a name sits in exactly one list"
        );
    }

    // What is not a picture is another kind's, however an engine could rasterise it. The
    // layered document and the scanner's dataset stay the image converter's, a metafile is the
    // drawing layer's, and a page of rich text is the text kind's.
    for (name, kind) in [
        ("picture.xcf", PreviewType::Magick),
        ("scan.dcm", PreviewType::Magick),
        ("picture.miff", PreviewType::Magick),
        ("still.jp2", PreviewType::Magick),
        ("drawing.emf", PreviewType::Vector),
        ("drawing.wmf", PreviewType::Vector),
    ] {
        let path = Path::new(name);

        assert_eq!(
            kind_of_name(path, &config),
            Some(kind),
            "`{name}` is the {kind:?} kind's"
        );
        assert!(
            !named_as(path, &config, PreviewType::Images),
            "`{name}` is not a picture of the master list"
        );
    }

    // And a page of rich text is the text kind's: it is drawn as a page rather than shown as a
    // picture, and a name that is the text list's today is not to be moved by a later change.
    assert_eq!(
        kind_of_name(Path::new("page.rtf"), &config),
        Some(PreviewType::Text),
        "a `.rtf` is a page of text"
    );

    // The two readers of the picture kind, and the one of the converter's, are what the chains say
    // as well: the pictures are read here first and developed by the engine only where that read
    // answered nothing.
    assert_eq!(
        chain(PreviewType::Images),
        &[Reader::Native, Reader::ImageMagick],
        "a picture is the readers of this app first"
    );
    assert_eq!(
        chain(PreviewType::Magick),
        &[Reader::ImageMagick],
        "the engine's own names have no reader in front of it"
    );

    // A model is the kind whose only reader is an engine, and whose list is the one that says so:
    // a `.stl` is a model by its name alone, and it is no other kind's read off the same name.
    let model = Path::new("model.stl");
    assert_eq!(
        kind_of_name(model, &config),
        Some(PreviewType::Models),
        "a `.stl` is a model"
    );
    assert_eq!(
        chain(PreviewType::Models),
        &[Reader::F3d],
        "and what draws one is the installed renderer or nothing"
    );
    assert!(named_as(model, &config, PreviewType::Models));
    for other in [
        PreviewType::Images,
        PreviewType::Vector,
        PreviewType::Design,
        PreviewType::Magick,
    ] {
        assert!(
            !named_as(model, &config, other),
            "a model is not the {other:?} kind's as well: a name sits in exactly one list"
        );
    }
}

/// **A model's name is half of the claim and the file's own front is the other half.** A `.bin`
/// is a name the `[f3d]` list carries and a name half this machine's formats share, so a file of
/// that name which is not a point cloud is not a model at all: it is answered with no kind, which
/// starts nothing, rather than with the kind that would launch the viewer for it (see
/// `model_formats`). A mesh is still claimed, because its front says what its name promised.
#[test]
fn a_model_is_the_kind_its_own_head_confirms_and_no_other() {
    let config = crate::config::config::AppConfig::default();
    let folder = std::env::temp_dir().join("rust-hover-preview-model-route");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let junk = folder.join("not-a-cloud.bin");
    std::fs::write(&junk, [0x9Au8, 0x7F, 0x01, 0x00].repeat(64)).expect("a file of junk");
    assert_eq!(
        kind_of_name(&junk, &config),
        Some(PreviewType::Models),
        "the name is still a model's: that half of the question has not changed"
    );
    assert_ne!(
        kind_of_with_facts(&junk, &config, None),
        Some(PreviewType::Models),
        "and the file is not one, so the viewer is never started for it"
    );
    let readers = kind_of_with_facts(&junk, &config, None)
        .map_or_else(Vec::new, |kind| readers_for(kind, &junk));
    assert!(
        !readers.contains(&Reader::F3d),
        "nothing of that route asks the viewer about a file that is not a model"
    );

    // 56 bytes, version 20020715, identifier 970401: the header of a real TerraScan file.
    let cloud = folder.join("a-cloud.bin");
    let mut head = b"8\x00\x00\x00\xeb\x7d\x31\x01\xa1\xce\x0e\x00".to_vec();
    head.extend_from_slice(&[0u8; 64]);
    std::fs::write(&cloud, &head).expect("a point cloud");
    assert_eq!(
        kind_of_with_facts(&cloud, &config, None),
        Some(PreviewType::Models),
        "a file whose header carries PDAL's version and identifier is the point cloud its name promised"
    );

    let mesh = folder.join("a-mesh.stl");
    std::fs::write(&mesh, b"solid cube\nendsolid cube\n").expect("an ascii mesh");
    assert_eq!(
        kind_of_with_facts(&mesh, &config, None),
        Some(PreviewType::Models),
        "an ascii mesh is a model by its name and by its front together"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// **A drawing is the one file two kinds can draw, and what settles it is the geometry it holds.**
/// The viewer reads only the 3D half of the DXF format — a `3DFACE`, and a `POLYLINE` whose flag
/// word says it is a 3D one, a mesh or a polyface — while the render engine beside it imports the
/// 2D drafting sheet whole. So a file of the name is claimed by its own group codes rather than by
/// the name, which is also why a `3DFACE` written as a label is still a sheet (see
/// `model_formats::dxf`), and the one thing beyond the file that the answer needs is the machine:
/// with no viewer here to draw it, the file stays the drawing it always was.
#[test]
fn a_drawing_is_the_kind_of_geometry_its_own_group_codes_name() {
    // A drawing of one entity list, in the sections a real file carries in front of it.
    fn drawing(entities: &str) -> String {
        format!(
            "0\nSECTION\n2\nHEADER\n9\n$ACADVER\n1\nAC1015\n0\nENDSEC\n\
             0\nSECTION\n2\nENTITIES\n{entities}0\nENDSEC\n0\nEOF\n"
        )
    }

    let folder = std::env::temp_dir().join("rust-hover-preview-dxf-route");
    std::fs::create_dir_all(&folder).expect("a test folder");

    // A machine with the viewer on it, without one being installed: the setting names a file, and
    // what the claim asks is whether there is one to run (`f3d_render::available_in`), which
    // nothing in the routing ever begins.
    let program = folder.join("f3d.exe");
    std::fs::write(&program, b"not a viewer, and never begun").expect("a program");
    let config = AppConfig {
        f3d_path: program.to_string_lossy().into_owned(),
        ..Default::default()
    };
    assert!(
        f3d_render::available_in(&config),
        "the faked machine is one with the viewer on it"
    );

    let written = |name: &str, body: &str| {
        let path = folder.join(name);
        std::fs::write(&path, body).expect("a drawing");

        path
    };

    // The geometry the viewer imports, an entity at a time, so a name that stopped being the
    // viewer's says which one did.
    for (name, entities) in [
        ("a-face.dxf", "0\n3DFACE\n8\n0\n10\n0\n11\n1\n"),
        ("a-3d-polyline.dxf", "0\nPOLYLINE\n66\n1\n70\n8\n"),
        ("a-mesh.dxf", "0\nPOLYLINE\n66\n1\n70\n16\n"),
        ("a-polyface.dxf", "0\nPOLYLINE\n66\n1\n70\n64\n"),
        // Mixed geometry and annotations is still the viewer's: the labels are dropped rather
        // than the scene.
        (
            "a-labelled-solid.dxf",
            "0\nMTEXT\n1\nplate 3\n0\n3DFACE\n8\n0\n0\nDIMENSION\n70\n32\n",
        ),
    ] {
        let path = written(name, &drawing(entities));

        assert_eq!(
            kind_of_with_facts(&path, &config, None),
            Some(PreviewType::Models),
            "`{name}` holds geometry the viewer reads, so it is a model"
        );
        assert_eq!(
            chain(PreviewType::Models),
            &[Reader::F3d],
            "and what draws a model is the viewer or nothing"
        );
    }

    for (name, entities) in [
        (
            "a-sheet.dxf",
            "0\nLWPOLYLINE\n90\n4\n70\n0\n0\nARC\n10\n0\n0\nTEXT\n1\nplate 3\n",
        ),
        // The ACIS bodies, which the viewer's importer cannot read at all.
        ("an-acis-body.dxf", "0\n3DSOLID\n70\n1\n0\nREGION\n70\n1\n"),
        // A name written as a value, which is a label rather than an entity.
        ("a-labelled-sheet.dxf", "0\nTEXT\n8\n3DFACE\n1\n3DFACE\n"),
        // A 2D polyline: the flags say the entity the viewer would read nothing of.
        ("a-flat-polyline.dxf", "0\nPOLYLINE\n70\n1\n"),
        // Past the bound the scan is bounded to, so the answer is the sheet's.
        (
            "a-sheet-past-the-bound.dxf",
            &format!(
                "{}{}",
                "0\nTEXT\n1\nfiller\n".repeat(40_000),
                "0\n3DFACE\n8\n0\n"
            ),
        ),
    ] {
        let path = written(name, &drawing(entities));

        assert_eq!(
            kind_of_with_facts(&path, &config, None),
            Some(PreviewType::Libre),
            "`{name}` holds nothing the viewer draws, so it is the render engine's sheet"
        );
        assert_eq!(
            chain(PreviewType::Libre),
            &[Reader::LibreOffice],
            "and what draws a sheet is the render engine"
        );
    }

    // The binary spelling is the render engine's whatever it holds: the viewer's importer reads
    // the ASCII one only.
    let binary = written(
        "a-binary-face.dxf",
        "AutoCAD Binary DXF\r\n\u{1a}\u{0}0\u{0}3\u{0}D\u{0}F\u{0}A\u{0}C\u{0}E\u{0}",
    );
    assert_eq!(
        kind_of_with_facts(&binary, &config, None),
        Some(PreviewType::Libre),
        "a drawing written as binary is the render engine's"
    );

    // And the same file through the whole of a route: what its bytes say is no opinion, since the
    // name it carries is the one the signature named, so the kind the loader draws it as is the
    // router's — which is what a 3D drawing reaching the viewer depends on. A content tier that
    // answered the render engine's kind here would hand the file to the engine that draws nothing
    // of it (see `content_type::classify`).
    let mesh = folder.join("a-face.dxf");
    let probe = crate::formats::content_type::Probe::read(&mesh);
    let route = resolve(&mesh, &config, &probe);
    assert_eq!(
        route.content,
        crate::formats::content_type::Content::Unknown,
        "a drawing under its own name is the router's to settle"
    );
    assert_eq!(route.named, Some(PreviewType::Models));

    // And the machine's half: with no viewer installed there is nothing to claim the scene, so
    // the drawing is the sheet's — which is the answer it had before the viewer was asked at all.
    let mesh = folder.join("a-face.dxf");
    assert_eq!(
        kind_of_with_facts(&mesh, &AppConfig::default(), None),
        if f3d_render::available() {
            Some(PreviewType::Models)
        } else {
            Some(PreviewType::Libre)
        },
        "a drawing of a mesh is the viewer's where one is installed and the sheet's where none is"
    );

    // And the two settings that keep it a sheet: the tray's `3D` gate, which is a preview nobody
    // asked for, and a list that no longer names the file — a machine with the viewer on it either
    // way.
    for closed in [
        AppConfig {
            f3d_preview_enabled: false,
            ..config.clone()
        },
        AppConfig {
            f3d_extensions: Vec::new(),
            ..config.clone()
        },
    ] {
        assert_eq!(
            kind_of_with_facts(&mesh, &closed, None),
            Some(PreviewType::Libre),
            "a drawing the viewer will not be asked to draw is the sheet's"
        );
    }

    // A name alone is answered without the file being read, which is the form `content_type`
    // asks: the 2D/3D split is the body's question, and a name has no body.
    assert_eq!(
        kind_of_name(&mesh, &config),
        Some(PreviewType::Libre),
        "a name with no file read behind it is the render engine's"
    );

    let _ = std::fs::remove_dir_all(&folder);
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
        ("picture.xcf", PreviewType::Magick),
        ("photo.jpg", PreviewType::Images),
        ("shot.nef", PreviewType::Images),
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
        (PreviewType::Models, NavCategory::Design),
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
        PreviewType::Models,
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
        probe.f3d_preview_enabled = false;
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
        &[Reader::Native, Reader::ImageMagick],
        "a picture is the readers of this app first, and the image converter behind them for the \
         codec package the machine has not been given"
    );
    assert_eq!(
        chain(PreviewType::Magick),
        &[Reader::ImageMagick],
        "and a name of the engine's own has no reader in front of it: nothing here opens one"
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
    let config = AppConfig {
        text_dont_preview_secrets: false,
        ..Default::default()
    };

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

            // The kinds an engine's own window draws rather than a reader of this app's, so they
            // are the arms that answer `Nothing` — the only way a claimed kind has no hand to be
            // drawn by. It is the four half-kinds an engine draws *for* another kind, and the one
            // kind that has no reader here at all.
            let engine_kind = matches!(
                route.named,
                Some(PreviewType::Document)
                    | Some(PreviewType::Libre)
                    | Some(PreviewType::Calibre)
                    | Some(PreviewType::Magick)
                    | Some(PreviewType::Models)
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

#[test]
fn extensionless_files_route_to_text_when_enabled() {
    let mut config = AppConfig::default();
    let extensionless_path = Path::new("extensionless_sample");
    let txt_path = Path::new("notes.txt");

    assert!(!config.text_show_extensionless);
    assert!(!named_as(extensionless_path, &config, PreviewType::Text));
    assert_eq!(kind_of(extensionless_path, &config), None);
    assert!(named_as(txt_path, &config, PreviewType::Text));
    assert_eq!(kind_of(txt_path, &config), Some(PreviewType::Text));

    config.text_show_extensionless = true;
    assert!(named_as(extensionless_path, &config, PreviewType::Text));
    assert_eq!(
        kind_of(extensionless_path, &config),
        Some(PreviewType::Text)
    );
    assert!(named_as(txt_path, &config, PreviewType::Text));
    assert_eq!(kind_of(txt_path, &config), Some(PreviewType::Text));

    config.text_show_extensionless = false;
    assert!(!named_as(extensionless_path, &config, PreviewType::Text));
    assert_eq!(kind_of(extensionless_path, &config), None);
    assert!(named_as(txt_path, &config, PreviewType::Text));
    assert_eq!(kind_of(txt_path, &config), Some(PreviewType::Text));
}

#[test]
fn secret_files_are_not_previewed_when_enabled_and_previewed_when_disabled() {
    let mut config = AppConfig::default();
    let env_file = Path::new(".env");
    let env_local = Path::new(".env.local");
    let env_example = Path::new(".env.example");

    // Enabled by default: secrets are blocked
    assert!(config.text_dont_preview_secrets);
    assert!(!named_as(env_file, &config, PreviewType::Text));
    assert_eq!(kind_of(env_file, &config), None);
    assert!(!named_as(env_local, &config, PreviewType::Text));
    assert_eq!(kind_of(env_local, &config), None);

    // .env.example is an exempt template, so it is previewed
    assert!(named_as(env_example, &config, PreviewType::Text));
    assert_eq!(kind_of(env_example, &config), Some(PreviewType::Text));

    // When disabled: secrets in text lists can be previewed
    config.text_dont_preview_secrets = false;
    assert!(named_as(env_file, &config, PreviewType::Text));
    assert_eq!(kind_of(env_file, &config), Some(PreviewType::Text));
    assert!(named_as(env_local, &config, PreviewType::Text));
    assert_eq!(kind_of(env_local, &config), Some(PreviewType::Text));
}
