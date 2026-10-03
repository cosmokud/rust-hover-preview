use super::*;

/// A file of the bytes given, under a name, in a folder of its own.
fn sample(folder: &str, name: &str, bytes: &[u8]) -> PathBuf {
    let folder = std::env::temp_dir().join(format!("rust-hover-preview-head-{folder}"));
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join(name);
    std::fs::write(&path, bytes).expect("a test file");

    path
}

/// The eight bytes every PNG opens with.
pub(crate) fn png_open() -> Vec<u8> {
    vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]
}

/// A chunk as a PNG writes one: its length, its type, its body and its CRC.
pub(crate) fn png_chunk(kind: &[u8; 4], body: &[u8]) -> Vec<u8> {
    let mut chunk = (body.len() as u32).to_be_bytes().to_vec();
    chunk.extend_from_slice(kind);
    chunk.extend_from_slice(body);
    chunk.extend_from_slice(&[0, 0, 0, 0]);

    chunk
}

/// A GIF's header and logical screen descriptor, with no global colour table.
pub(crate) fn gif_open() -> Vec<u8> {
    let mut gif = b"GIF89a".to_vec();
    gif.extend_from_slice(&[0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00]);

    gif
}

/// A GIF image block: its descriptor, its code size and one run of sub-blocks.
pub(crate) fn gif_frame() -> Vec<u8> {
    let mut frame = vec![0x2C];
    frame.extend_from_slice(&[0u8; 9]);
    frame.push(0x02);
    frame.extend_from_slice(&[0x03, 0x00, 0x00, 0x00, 0x00]);

    frame
}

/// A WebP's container header and one chunk.
fn webp_chunk(kind: &[u8; 4], body: &[u8]) -> Vec<u8> {
    let mut chunk = kind.to_vec();
    chunk.extend_from_slice(&(body.len() as u32).to_le_bytes());
    chunk.extend_from_slice(body);

    chunk
}

/// A TIFF whose first image directory sits at eight and holds no entries, with the four
/// bytes after them pointing at a second directory or at nothing.
fn tiff(next_page: bool) -> Vec<u8> {
    let mut tiff = b"II*\0".to_vec();
    tiff.extend_from_slice(&8u32.to_le_bytes());
    tiff.extend_from_slice(&0u16.to_le_bytes());
    tiff.extend_from_slice(&if next_page { 12u32 } else { 0u32 }.to_le_bytes());

    tiff
}

#[test]
fn a_still_picture_and_one_that_moves_are_told_apart_by_their_own_bytes() {
    let still_png = sample(
        "still-png",
        "picture.png",
        &[
            png_open(),
            png_chunk(b"IHDR", &[0u8; 13]),
            png_chunk(b"IDAT", &[0]),
        ]
        .concat(),
    );
    let animated_png = sample(
        "animated-png",
        "picture.png",
        &[
            png_open(),
            png_chunk(b"IHDR", &[0u8; 13]),
            png_chunk(b"acTL", &[0u8; 8]),
            png_chunk(b"IDAT", &[0]),
        ]
        .concat(),
    );

    assert_eq!(
        picture_nature(&still_png).map(|nature| nature.form),
        Some(PictureForm::Still)
    );
    assert_eq!(
        picture_nature(&animated_png),
        Some(nature(PictureForm::Plays(PictureFamily::Apng), true, false)),
        "the same `.png` name, and its own `acTL` chunk is what says which it is"
    );

    let still_gif = sample(
        "still-gif",
        "picture.gif",
        &[gif_open(), gif_frame(), vec![0x3B]].concat(),
    );
    let animated_gif = sample(
        "animated-gif",
        "picture.gif",
        &[gif_open(), gif_frame(), gif_frame(), vec![0x3B]].concat(),
    );

    assert_eq!(
        picture_nature(&still_gif).map(|nature| nature.form),
        Some(PictureForm::Still)
    );
    assert_eq!(
        picture_nature(&animated_gif).map(|nature| nature.form),
        Some(PictureForm::Plays(PictureFamily::Gif))
    );

    let mut still_webp = b"RIFF".to_vec();
    still_webp.extend_from_slice(&0u32.to_le_bytes());
    still_webp.extend_from_slice(b"WEBP");
    still_webp.extend_from_slice(&webp_chunk(b"VP8 ", &[0; 8]));

    let mut animated_webp = b"RIFF".to_vec();
    animated_webp.extend_from_slice(&0u32.to_le_bytes());
    animated_webp.extend_from_slice(b"WEBP");
    animated_webp.extend_from_slice(&webp_chunk(b"VP8X", &[0; 10]));
    animated_webp.extend_from_slice(&webp_chunk(b"ANIM", &[0; 6]));

    assert_eq!(
        picture_nature(&sample("still-webp", "picture.webp", &still_webp))
            .map(|nature| nature.form),
        Some(PictureForm::Still)
    );
    assert_eq!(
        picture_nature(&sample("animated-webp", "picture.webp", &animated_webp)),
        Some(nature(PictureForm::Plays(PictureFamily::Webp), true, false))
    );
}

/// An ISO base media sequence: the brand at the front is what says whether it is one
/// picture or a sequence of them, and the whole of the brand list is read — not only
/// the major one, which is the half that made an animated AVIF look like a still one.
#[test]
fn an_iso_base_media_sequence_is_told_apart_by_the_brands_it_declares() {
    // `avis` in front, which is the sequence form of AVIF.
    let mut avif_sequence = b"\x00\x00\x00\x20".to_vec();
    avif_sequence.extend_from_slice(b"ftypavis");
    avif_sequence.extend_from_slice(&[0u8; 16]);

    // `mif1`, the generic HEIF brand a HEIC burst is usually written under, with `heic`
    // behind it as a compatible brand. `mif1` on its own is *not* a sequence: it is the
    // brand an ordinary single-image HEIC is written under too, so it is the still case
    // as much as the moving one, and is settled as a picture.
    let mut heic_single_under_mif1 = b"\x00\x00\x00\x20".to_vec();
    heic_single_under_mif1.extend_from_slice(b"ftypmif1");
    heic_single_under_mif1.extend_from_slice(&[0u8; 8]);
    heic_single_under_mif1.extend_from_slice(b"heic");
    heic_single_under_mif1.extend_from_slice(&[0u8; 8]);

    // A HEIF sequence, which is `msf1` — the brand that does say so.
    let mut heif_sequence = b"\x00\x00\x00\x20".to_vec();
    heif_sequence.extend_from_slice(b"ftypmsf1");
    heif_sequence.extend_from_slice(&[0u8; 16]);

    // The ordinary way an animated AVIF is written: a still major brand, with the
    // sequence brand behind it. Reading only the front four bytes misses this one.
    let mut avis_behind_avif = b"\x00\x00\x00\x28".to_vec();
    avis_behind_avif.extend_from_slice(b"ftypavif");
    avis_behind_avif.extend_from_slice(&[0u8; 8]);
    avis_behind_avif.extend_from_slice(b"avis");
    avis_behind_avif.extend_from_slice(&[0u8; 8]);

    // A still AVIF: `avif` in front and no sequence brand anywhere behind it.
    let mut single = b"\x00\x00\x00\x20".to_vec();
    single.extend_from_slice(b"ftypavif");
    single.extend_from_slice(&[0u8; 16]);

    // A still HEIC, which is the ordinary `.heic` a camera writes.
    let mut heic_single = b"\x00\x00\x00\x20".to_vec();
    heic_single.extend_from_slice(b"ftypheic");
    heic_single.extend_from_slice(&[0u8; 16]);

    assert_eq!(
        picture_nature(&sample("avis", "picture.avif", &avif_sequence)),
        Some(nature(PictureForm::Plays(PictureFamily::Heif), true, false)),
        "an `avis` brand is a sequence in time, and it is played"
    );
    assert_eq!(
        picture_nature(&sample("msf1", "picture.heic", &heif_sequence)),
        Some(nature(PictureForm::Plays(PictureFamily::Heif), true, false)),
        "`msf1` is the HEIF sequence brand, and it is a sequence whatever it codes with"
    );
    assert_eq!(
        picture_nature(&sample(
            "avis-behind-avif",
            "picture.avif",
            &avis_behind_avif
        )),
        Some(nature(PictureForm::Plays(PictureFamily::Heif), true, false)),
        "an `avis` behind an `avif` is still a sequence: the major brand is not the only brand"
    );
    assert_eq!(
        picture_nature(&sample("avif", "picture.avif", &single)),
        Some(nature(PictureForm::Still, false, false))
    );
    assert_eq!(
        picture_nature(&sample("heic", "picture.heic", &heic_single)),
        Some(nature(PictureForm::Still, false, false)),
        "a `heic` with no sequence brand behind it is the still picture a camera writes"
    );
    assert_eq!(
        picture_nature(&sample("mif1", "picture.heic", &heic_single_under_mif1)),
        Some(nature(PictureForm::Still, false, false)),
        "`mif1` is the generic HEIF *image* brand, and a single-image HEIC is written \
         under it — treating it as a sequence would give every camera photo the \
         animated scale"
    );

    // A codec brand behind a container brand says what the file was made with, not
    // what it is: a HEIF that lists `av01` is a HEIC and must still be answered as one.
    let mut heic_coded_with_av1 = b"\x00\x00\x00\x20".to_vec();
    heic_coded_with_av1.extend_from_slice(b"ftypheic");
    heic_coded_with_av1.extend_from_slice(&[0u8; 8]);
    heic_coded_with_av1.extend_from_slice(b"av01");
    heic_coded_with_av1.extend_from_slice(&[0u8; 8]);

    assert_eq!(
        picture_nature(&sample("heic-av01", "picture.heic", &heic_coded_with_av1)),
        Some(nature(PictureForm::Still, false, false)),
        "an `av01` behind a `heic` is the codec it was made with, not an AVIF"
    );

    let mng = sample("mng", "picture.mng", &[0x8A, b'M', b'N', b'G', 0, 0, 0, 0]);
    assert_eq!(
        picture_nature(&mng),
        Some(nature(PictureForm::Unplayable, true, false))
    );
}

/// A JPEG XL is recognised in both the forms it arrives in — a naked codestream and a
/// container — from its own signature, whichever it is.
///
/// Whether one *moves* is a different question, and is not answered here: the anim flag
/// is in the codestream's image header, behind the container box, so it is the decoder
/// that is asked (`jxl_image::is_animated`). What this test holds is that the signature
/// is recognised at all — which it was not, before: a `.jxl` fell through the whole of
/// this module and was left to the tables.
#[test]
fn a_jpeg_xl_is_recognised_in_both_the_forms_it_arrives_in() {
    let codestream = sample(
        "jxl-codestream",
        "picture.jxl",
        &[0xFF, 0x0A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00],
    );
    let container = sample(
        "jxl-container",
        "picture.jxl",
        &[b"\x00\x00\x00\x0CJXL \r\n\x87\n".to_vec(), vec![0u8; 16]].concat(),
    );

    // Neither of the two above carries a codestream, so neither is a sequence — but both
    // are JPEG XL files, and both are answered as the still pictures they are rather
    // than as a front that settles nothing.
    assert_eq!(
        picture_nature(&codestream),
        Some(nature(PictureForm::Still, false, false)),
        "a naked codestream is a JPEG XL"
    );
    assert_eq!(
        picture_nature(&container),
        Some(nature(PictureForm::Still, false, false)),
        "the container form is a JPEG XL too"
    );

    // A file that is none of these is still a front that settles nothing.
    assert_eq!(
        picture_nature(&sample("jxl-none", "picture.jxl", b"not a jpeg xl")),
        None
    );
}

#[test]
fn more_than_one_page_is_not_a_thing_that_moves() {
    let single = sample("tiff-one", "page.tif", &tiff(false));
    let paged = sample("tiff-many", "page.tif", &tiff(true));

    assert_eq!(
        picture_nature(&single).map(|nature| nature.form),
        Some(PictureForm::Still)
    );
    assert_eq!(
        picture_nature(&paged).map(|nature| nature.form),
        Some(PictureForm::Paged)
    );

    for path in [&single, &paged] {
        assert!(
            picture_nature(path).is_some_and(|nature| !nature.moves),
            "a TIFF with more pages than one is pages, not an animation"
        );
    }

    let one_icon = sample("ico-one", "icon.ico", &[0x00, 0x00, 0x01, 0x00, 0x01, 0x00]);
    let many_icons = sample(
        "ico-many",
        "icon.ico",
        &[0x00, 0x00, 0x01, 0x00, 0x03, 0x00],
    );

    assert_eq!(
        picture_nature(&one_icon).map(|nature| nature.form),
        Some(PictureForm::Still)
    );
    assert_eq!(
        picture_nature(&many_icons).map(|nature| nature.form),
        Some(PictureForm::Paged)
    );
}

#[test]
fn a_tiff_settles_nothing_on_its_own_because_a_raw_is_written_in_one() {
    let raw = sample("tiff-raw", "shot.cr2", &tiff(false));

    assert!(
        picture_nature(&raw).is_some_and(|nature| nature.container_of_a_raw),
        "the front of a raw is the front of a TIFF, and which reader is asked about \
         one is the byte tables' question"
    );
}

#[test]
fn the_front_is_read_no_further_where_it_settles_the_question() {
    let png = sample(
        "front",
        "picture.png",
        &[
            png_open(),
            png_chunk(b"IHDR", &[0u8; 13]),
            png_chunk(b"IDAT", &[0]),
        ]
        .concat(),
    );

    let front = of(&png).expect("a head");
    assert_eq!(front.bytes().len(), SHORT_BYTES);
    assert!(!front.complete(), "the rest of the window was not read");

    let whole = full(&png).expect("a head");
    assert!(whole.complete());
    assert!(
        whole.bytes().len() > SHORT_BYTES,
        "the whole window is what the tables are asked about"
    );

    // A file whose front is nothing this module knows is read to the whole window,
    // because what it is may be further in than the front.
    let unknown = sample("unknown", "data.bin", &[0x11u8; 64]);
    let head = of(&unknown).expect("a head");
    assert!(head.complete());
    assert_eq!(head.bytes().len(), 64);
}

#[test]
fn both_windows_answer_the_same_question_the_same_way() {
    let png = sample(
        "agree-png",
        "picture.png",
        &[
            png_open(),
            png_chunk(b"IHDR", &[0u8; 13]),
            png_chunk(b"acTL", &[0u8; 8]),
            png_chunk(b"IDAT", &[0u8; 512]),
        ]
        .concat(),
    );
    let gif = sample(
        "agree-gif",
        "picture.gif",
        &[gif_open(), gif_frame(), gif_frame(), vec![0x3B]].concat(),
    );
    let tiff = sample("agree-tiff", "page.tif", &tiff(true));

    for path in [&png, &gif, &tiff] {
        let front = of(path).expect("a head");
        let whole = full(path).expect("a head");

        assert_eq!(
            front.nature(),
            whole.nature(),
            "reading on from the front never changes what the front said"
        );
        assert_eq!(front.front(), whole.front());
    }
}
