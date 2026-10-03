use super::*;

/// A name and a content that disagree are answered with the kind the content belongs
/// to, which is what a file is handed to an engine by.
#[test]
fn a_disagreement_is_answered_with_the_kind_the_content_belongs_to() {
    assert_eq!(
        classified("film.docx", b"\x00\x00\x00\x20ftypisom"),
        Content::Kind(PreviewType::Videos),
        "an MP4 under a document's name is a video"
    );
    assert_eq!(
        classified("drawing.cdr", b"\x89PNG\x0D\x0A\x1A\x0A"),
        Content::Kind(PreviewType::Images),
        "a picture under a drawing's name is a picture"
    );
    assert_eq!(
        classified("sheet.wpd", b"GIF89a"),
        Content::Kind(PreviewType::Images),
        "and a GIF is a picture wherever it is found"
    );
    // A transport stream is one of the formats only this app's table names: the
    // common table has no entry for one at all. What such a file looks like is a sync
    // byte every packet size apart, which is the probe the text lists share.
    let transport_stream = {
        let mut probe = vec![0u8; 188 * 4];
        for packet in 0..4 {
            probe[packet * 188] = 0x47;
        }
        probe
    };
    assert_eq!(
        classified("film.docx", &transport_stream),
        Content::Kind(PreviewType::Videos),
        "a transport stream under a document's name is a video"
    );
    assert_eq!(
        classified("letter.docx", b"%!PS-Adobe-3.0"),
        Content::Kind(PreviewType::Vector),
        "and PostScript is the drawing the vector list reads"
    );
}

/// Where the content and the name agree there is nothing to override: the file has
/// the kind it was always going to have, whether the name it carries is the first the
/// format is known by or one of the others.
#[test]
fn agreement_is_no_opinion() {
    assert_eq!(
        classified("shot.png", b"\x89PNG\x0D\x0A\x1A\x0A"),
        Content::Unknown
    );
    assert_eq!(
        classified("film.mkv", &[0x1A, 0x45, 0xDF, 0xA3]),
        Content::Unknown
    );
    assert_eq!(
        classified("film.mp4", b"\x00\x00\x00\x20ftypisom"),
        Content::Unknown
    );
    assert_eq!(
        classified("picture.mjpeg", &[0xFF, 0xD8, 0xFF, 0xE0]),
        Content::Unknown,
        "a motion JPEG is a video by name as well as a picture by content"
    );
}

/// A box is not an answer, which is the half of the signature tables that is
/// deliberately not in them: an OpenDocument and an Office package are both zips, and
/// neither table answers for one, so the name the file was given is what decides —
/// which for a name the third table holds is the engine that draws it, and for every
/// other name is the list that claims it.
#[test]
fn a_box_is_left_to_the_name() {
    for name in ["letter.odt", "report.docx", "bundle.zip"] {
        assert_eq!(
            classified(name, b"PK\x03\x04\x14\x00\x00\x00"),
            Content::Unknown,
            "`{name}` is a zip, and a zip is not a kind"
        );
    }

    // An OLE compound file — every `.doc`, `.xls` and `.ppt` — is a box of the same
    // kind, and the name is what is answered with.
    assert_eq!(
        classified("letter.doc", b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"),
        Content::Unknown
    );
}

/// A format whose head is nothing a signature names is answered by the name it is
/// written under, which is what the table of names is for: the names no probe can be
/// asked about, and the files whose name is all there is to route them by.
#[test]
fn a_format_no_signature_names_is_answered_by_the_name_it_carries() {
    // A TiVo stream, whose chunk headers are a hundred and twenty-eight kilobytes
    // apart, and a Cineon file, whose demuxer is gone: two names no head can settle.
    assert_eq!(
        classified("recording.ty", b"\x00"),
        Content::Kind(PreviewType::Videos)
    );
    assert_eq!(
        classified("clip.ty+", b"\x00"),
        Content::Kind(PreviewType::Videos),
        "a name the lists carry and no signature does is answered as it is written"
    );
    assert_eq!(
        classified("film.cin", b"\x01\x02\x03\x04"),
        Content::Kind(PreviewType::Videos)
    );

    // A flat OpenDocument is a document whose signature is its own markup, and nothing
    // is asked of one: the list the name is written for is what decides.
    assert_eq!(
        classified("sheet.fodt", b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"),
        Content::Kind(PreviewType::Libre)
    );

    // A WordPerfect document that writes its header is the render engine's by its own
    // bytes. One that writes no header is a file no table here has an opinion about,
    // which is the state every name in the table below is in.
    assert_eq!(
        classified("letter.docx", b"\xFFWPC\x00\x00\x00\x00\x01\x0A"),
        Content::Kind(PreviewType::Libre),
        "a WordPerfect document under a document's name is the engine's"
    );
    assert_eq!(
        classified("letter.wpd", b"nothing in here is a signature"),
        Content::Unknown
    );

    // A name whose format is a package is a zip in the bytes and a document in the
    // name — an iWork deck is one — and the table answers with the engine's kind,
    // which is what the list it is written for would have said.
    assert_eq!(
        classified("deck.key", b"PK\x03\x04\x14\x00\x00\x00"),
        Content::Kind(PreviewType::Libre)
    );
}

/// The packages that declare their own type are the one shape of container this module
/// answers for, and what it reads is that declaration — which is what keeps the rule
/// that a container is not a kind, because the type is the document's own answer rather
/// than the box's.
#[test]
fn a_package_is_answered_by_the_type_it_declares() {
    assert_eq!(
        classified(
            "film.docx",
            &declared_package("application/vnd.oasis.opendocument.graphics")
        ),
        Content::Kind(PreviewType::Libre),
        "a drawing is the render engine's, whatever the file is called"
    );
    assert_eq!(
        classified(
            "film.docx",
            &declared_package("application/vnd.sun.xml.writer")
        ),
        Content::Kind(PreviewType::Libre),
        "and so is a StarOffice document of the XML generation"
    );
    assert_eq!(
        classified("photo.png", &declared_package("application/x-krita")),
        Content::Kind(PreviewType::Design),
        "while a Krita project is a design document"
    );

    // A template's type begins the name of the type it is a template of, and the record
    // that follows the declaration is what settles which of the two a file carries: a
    // `…graphics-template` package is not answered as `…graphics`.
    assert!(!is_declared_package(
        &declared_package("application/vnd.oasis.opendocument.graphics-template"),
        b"application/vnd.oasis.opendocument.graphics"
    ));
    assert!(is_declared_package(
        &declared_package("application/vnd.oasis.opendocument.graphics-template"),
        b"application/vnd.oasis.opendocument.graphics-template"
    ));

    // And a zip that declares nothing is a zip: an Office package, an OpenDocument of
    // the version that says `mimetype` rather than storing it as the first entry, and
    // any other archive are one answer, which is the name.
    assert_eq!(
        classified("letter.docx", b"PK\x03\x04\x14\x00\x00\x00"),
        Content::Unknown
    );
}

/// The name `.pdb` is two formats, and only one of them is a document: the Palm OS
/// database the render engine's own filters read as an ebook, and the Microsoft
/// program database a compiler writes beside its binaries. The engine is asked about
/// one of them and never about the other.
#[test]
fn the_two_formats_one_name_holds_are_told_apart() {
    let ebook = palm_database("Huckleberry Finn", b"TEXt", b"REAd");
    assert_eq!(
        classified("book.pdb", &ebook),
        Content::Kind(PreviewType::Libre),
        "a Palm OS ebook is the document the engine draws"
    );

    // The program database is the one a developer's folders are full of, and it is no
    // kind of document at all: nothing is shown and no engine is started for it.
    let program_database = {
        let mut probe = b"Microsoft C/C++ MSF 7.00\r\n\x1aDS\x00\x00\x00".to_vec();
        probe.resize(1024, 0);
        probe
    };
    assert_eq!(
        classified("app.pdb", &program_database),
        Content::Foreign,
        "a program database starts no engine"
    );
    assert_eq!(
        classified(
            "app.pdb",
            b"Microsoft C/C++ program database 2.00\r\n\x1aJG"
        ),
        Content::Foreign,
        "and neither does one of the version before it"
    );

    // What the engine reads is the Palm document, and a file of the name that is
    // nothing of the sort is not left to the name to decide: it is a file this app has
    // no reader for, which is the answer to the program database as well.
    assert_eq!(
        classified("data.pdb", b"\x00\x01\x02\x03"),
        Content::Foreign
    );
    assert_eq!(
        classified("data.pdb", b"not a database of any kind"),
        Content::Foreign
    );
}

/// And a name the signature tables *do* answer is not in the table of names: `.ts` and
/// `.mts` are the two a text list shares, and a file of one of those names that is not
/// the transport stream is the TypeScript source it is named as — which is a question
/// only the bytes can settle.
#[test]
fn a_name_a_signature_answers_is_not_answered_by_its_name() {
    for (extension, _) in KIND_BY_NAME {
        for signature in SIGNATURES {
            assert!(
                !signature
                    .names
                    .iter()
                    .any(|name| name.eq_ignore_ascii_case(extension)),
                "`{extension}` is answered by a signature, so the name must not answer"
            );
        }
    }

    assert_eq!(
        classified("app.ts", b"const x: number = 1;\n"),
        Content::Unknown,
        "a TypeScript file is left to the text list it is named under"
    );
}

/// What is in the table is the engines' own lists, and the kind beside a name is
/// the kind the lists themselves give it: a video name is a video, a name of the
/// render engine's is the engine's, a name of the ebook engine's is the ebook engine's, and a
/// name no list carries is not in the table at all. It is what keeps the table from drifting
/// away from the lists — a name taken out of `libre_formats`, the way `swf` was, has to be
/// taken out of here with it — and what keeps `dif`, which two lists carry, answered in the
/// order every other question about a file is asked in.
#[test]
fn the_table_holds_the_names_the_lists_hold() {
    let config = AppConfig::default();

    for (extension, kind) in KIND_BY_NAME {
        let named = PathBuf::from(format!("content.{extension}"));

        if crate::formats::video_formats::claims_any_video_name(&named, &config) {
            assert_eq!(*kind, PreviewType::Videos, "`{extension}` is a video name");
            continue;
        }

        if crate::formats::lists::LIBRE.claims(&named, &config) {
            assert_eq!(
                *kind,
                PreviewType::Libre,
                "`{extension}` is the engine's name"
            );
            continue;
        }

        if crate::formats::lists::CALIBRE.claims(&named, &config) {
            assert_eq!(
                *kind,
                PreviewType::Calibre,
                "`{extension}` is the ebook engine's name"
            );
            continue;
        }

        panic!("`{extension}` is not a name any list carries");
    }
}

/// A format no kind of this app previews is a file with nothing to show, and no engine
/// is started for it.
#[test]
fn a_format_no_kind_previews_is_answered_with_nothing() {
    let executable = {
        let mut probe = vec![0u8; 0x80];
        probe[0] = b'M';
        probe[1] = b'Z';
        probe[0x3C] = 0x40;
        probe[0x40..0x44].copy_from_slice(b"PE\x00\x00");
        probe
    };

    assert_eq!(
        classified("letter.docx", &executable),
        Content::Foreign,
        "an executable under a document's name starts no engine"
    );

    // A Unix program is one of those, and the answer for one is the same: nothing is
    // shown and nothing is started.
    let unix_program = {
        let mut probe = vec![0u8; 64];
        probe[..4].copy_from_slice(&[0x7F, b'E', b'L', b'F']);
        probe[4] = 0x02;
        probe[5] = 0x01;
        probe[6] = 0x01;
        probe
    };
    assert_eq!(
        classified("letter.docx", &unix_program),
        Content::Foreign,
        "and neither does a Unix one"
    );
    // And a sound, which is the one that changed its answer: an MP3 under a document's
    // name is the sound the sound list plays rather than a format with nothing to show,
    // because the kind exists now (see `audio_formats`). What still has no preview of its
    // own is a MIDI sequence, which neither engine this app has plays.
    assert_eq!(
        classified("letter.docx", b"ID3\x04\x00\x00\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Audio),
        "a song is not a document either — it is a song"
    );
    assert_eq!(
        classified("tune.docx", b"MThd\x00\x00\x00\x06\x00\x01"),
        Content::Foreign,
        "and a sequence of notes is neither: nothing here plays one"
    );
}

/// A front nothing in either table names, under a name the third table does not hold
/// either, is left to the name it has — which is the state a text file is in, and
/// every format this app reads for itself and confirms nothing about.
#[test]
fn what_has_no_signature_is_left_to_the_name_it_has() {
    assert_eq!(
        classified("stream.h264", b"\x00\x01\x02\x03 not a format at all"),
        Content::Unknown,
        "a `.h264` that is not one is answered by its name like any other file"
    );
    assert_eq!(
        classified("notes.txt", b"just some text\n"),
        Content::Unknown
    );
    assert_eq!(
        classified("drawing.svg", b"<svg xmlns=\"http://www.w3.org/2000/svg\">"),
        Content::Unknown,
        "a document whose signature is its own text is left to the lists"
    );
}

/// What the tables make of the files in a folder, asked the way a hover asks it —
/// through the reader, the setting and the cache rather than the rule alone. The
/// fixtures are the caller's: `RHP_CONTENT_PROBE` names a folder of real files, one of
/// every format the lists carry and a few named as another kind, which is the check
/// no table of its own can make.
#[test]
#[ignore = "reads the files named in RHP_CONTENT_PROBE"]
fn content_probe() {
    let folder = std::env::var("RHP_CONTENT_PROBE").expect("RHP_CONTENT_PROBE is not set");

    let mut paths: Vec<_> = std::fs::read_dir(&folder)
        .expect("the folder named by RHP_CONTENT_PROBE")
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .collect();
    paths.sort();

    let Ok(config) = crate::CONFIG.lock() else {
        return;
    };

    for path in paths {
        println!(
            "{:<16} {:?}",
            path.file_name().unwrap_or_default().to_string_lossy(),
            of(&path, &config)
        );
    }
}
