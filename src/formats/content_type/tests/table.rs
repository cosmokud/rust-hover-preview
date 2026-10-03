use super::*;

/// Every format the signature table carries is answered by its own head, under a name
/// that belongs to another kind. That is what the table is for: a hover meets files
/// that are not named for what they are, and the bytes are what settle it. The name in
/// each of these is one whose list would have sent the file somewhere else — to Word,
/// to the picture decoder, to the text preview — so what is asserted is that the
/// content wins.
#[test]
fn every_format_the_table_carries_is_answered_by_its_own_head() {
    // ---------------------------------------------------------------- pictures
    assert_eq!(
        classified("film.docx", b"DDS \x7C\x00\x00\x00"),
        Content::Kind(PreviewType::Images),
        "a DirectDraw Surface"
    );
    assert_eq!(
        classified("film.docx", &[0x76, 0x2F, 0x31, 0x01]),
        Content::Kind(PreviewType::Images),
        "an OpenEXR picture"
    );
    assert_eq!(
        classified("film.docx", b"#?RGBE\nFORMAT=32-bit_rle_rgbe\n"),
        Content::Kind(PreviewType::Images),
        "a Radiance picture of the older header"
    );
    assert_eq!(
        classified("film.docx", b"#?RADIANCE\nFORMAT=32-bit_rle_rgbe\n"),
        Content::Kind(PreviewType::Images),
        "and one of the newer"
    );
    assert_eq!(
        classified("film.docx", b"farbfeld\x00\x00\x00\x02\x00\x00\x00\x02"),
        Content::Kind(PreviewType::Images),
        "a farbfeld picture"
    );
    assert_eq!(
        classified("film.docx", b"qoif\x00\x00\x00\x02\x00\x00\x00\x02\x03\x00"),
        Content::Kind(PreviewType::Images),
        "a Quite OK Image"
    );
    assert_eq!(
        classified("film.docx", b"P5 2 2 255\n\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Images),
        "a Netpbm picture"
    );
    assert_eq!(
        classified(
            "film.docx",
            b"P7\nWIDTH 1\nHEIGHT 1\nDEPTH 3\nMAXVAL 255\nENDHDR\n"
        ),
        Content::Kind(PreviewType::Images),
        "a PAM picture, which names its own fields"
    );
    assert_eq!(
        classified("film.docx", &padded(2048, b"PCD_IPI")),
        Content::Kind(PreviewType::Libre),
        "a Photo CD image pac"
    );
    assert_eq!(
        classified(
            "film.docx",
            &[0x0A, 0x05, 0x01, 0x08, 0, 0, 0, 0, 9, 9, 9, 9]
        ),
        Content::Kind(PreviewType::Libre),
        "a PCX picture"
    );
    assert_eq!(
        classified("film.docx", &[0x59, 0xA6, 0x6A, 0x95]),
        Content::Kind(PreviewType::Libre),
        "a Sun raster"
    );
    assert_eq!(
        classified("film.docx", &padded(522, &[0x00, 0x11, 0x02, 0xFF])),
        Content::Kind(PreviewType::Libre),
        "a QuickDraw PICT, whose version operator is past the header it is drawn with"
    );
    // Two pictures the common table is what names, which this app's own table has
    // nothing to say about: all that is asked is that the answer it gives is used.
    assert_eq!(
        classified("film.docx", &[0xFF, 0x0A, 0x00]),
        Content::Kind(PreviewType::Images),
        "a JPEG XL picture"
    );
    assert_eq!(
        classified("film.docx", &declared_package("image/openraster")),
        Content::Kind(PreviewType::Design),
        "an OpenRaster project, named by the type it declares"
    );

    // ------------------------------------------------- the archives an engine lists
    assert_eq!(
        classified("film.docx", b"!<arch>\ndebian-binary   "),
        Content::Kind(PreviewType::Peazip),
        "a Unix archive, which is what a `.deb` and a `.udeb` are"
    );
    assert_eq!(
        classified("film.docx", &[0x60, 0xEA, 0x1A, 0x00, 0x00, 0x00]),
        Content::Kind(PreviewType::Peazip),
        "an ARJ archive"
    );
    assert_eq!(
        classified("film.docx", b"MSCF\x00\x00\x00\x00\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "a cabinet file"
    );
    assert_eq!(
        classified("film.docx", b"ITSF\x03\x00\x00\x00\x60\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "a compiled help file"
    );
    assert_eq!(
        classified("film.docx", b"07070100000000"),
        Content::Kind(PreviewType::Peazip),
        "a cpio archive, in its ASCII header"
    );
    assert_eq!(
        classified("film.docx", &[0xC7, 0x71, 0x00, 0x00, 0x00, 0x00]),
        Content::Kind(PreviewType::Peazip),
        "and in the older binary one, written either way round"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00-lh5-\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "an LHA archive"
    );
    assert_eq!(
        classified("film.docx", &[0xED, 0xAB, 0xEE, 0xDB, 0x03, 0x00]),
        Content::Kind(PreviewType::Peazip),
        "a Linux package"
    );
    assert_eq!(
        classified("film.docx", b"xar!\x00\x1C\x00\x01"),
        Content::Kind(PreviewType::Peazip),
        "an Apple archive, which a `.pkg` and a `.xip` are"
    );
    assert_eq!(
        classified("film.docx", b"MSWIM\x00\x00\x00\x00\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "a Windows image"
    );
    assert_eq!(
        classified("film.docx", b"sqsh\x02\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "a SquashFS image"
    );
    assert_eq!(
        classified("film.docx", &[0x1F, 0x9D, 0x90, 0x00]),
        Content::Kind(PreviewType::Peazip),
        "a `.z` of the older Unix compressor"
    );
    assert_eq!(
        classified("film.docx", b"BZh9\x31\x41\x59\x26\x53\x59"),
        Content::Kind(PreviewType::Peazip),
        "a bzip2 stream"
    );
    assert_eq!(
        classified("film.docx", &[0xFD, b'7', b'z', b'X', b'Z', 0x00, 0x00]),
        Content::Kind(PreviewType::Peazip),
        "an xz stream"
    );
    assert_eq!(
        classified("film.docx", &[0x28, 0xB5, 0x2F, 0xFD, 0x00]),
        Content::Kind(PreviewType::Peazip),
        "a zstd stream"
    );
    assert_eq!(
        classified("film.docx", b"koly\x00\x00\x00\x04\x00\x00\x02\x00"),
        Content::Kind(PreviewType::Peazip),
        "a macOS disk image"
    );
    assert_eq!(
        classified("film.docx", b"conectix\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "a Virtual PC image"
    );
    assert_eq!(
        classified("film.docx", b"vhdxfile\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "a Hyper-V image"
    );
    assert_eq!(
        classified("film.docx", b"KDMV\x01\x00\x00\x00"),
        Content::Kind(PreviewType::Peazip),
        "a VMware image"
    );
    assert_eq!(
        classified("film.docx", b"QFI\xFB\x00\x00\x00\x02"),
        Content::Kind(PreviewType::Peazip),
        "a QEMU image"
    );
    assert_eq!(
        classified("film.docx", &at_offset(64, &[0x7F, 0x10, 0xDA, 0xBE])),
        Content::Kind(PreviewType::Peazip),
        "a VirtualBox image, whose magic sits sixty-four bytes in"
    );
    assert_eq!(
        classified("film.docx", &at_offset(32, b"NXSB")),
        Content::Kind(PreviewType::Peazip),
        "an Apple file system image"
    );
    assert_eq!(
        classified("film.docx", &at_offset(1024, b"BDH+")),
        Content::Kind(PreviewType::Peazip),
        "a Macintosh volume"
    );
    assert_eq!(
        classified("film.docx", &at_offset(16, b"Compressed ROMFS")),
        Content::Kind(PreviewType::Peazip),
        "a compressed file system of the older header"
    );

    // ---------------------------------------------------------------- drawings
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 88];
            probe[..4].copy_from_slice(&[0x01, 0x00, 0x00, 0x00]);
            probe[4..8].copy_from_slice(&88u32.to_le_bytes());
            probe[40..44].copy_from_slice(b" EMF");
            probe[44..48].copy_from_slice(&[0x00, 0x00, 0x01, 0x00]);
            probe
        }),
        Content::Kind(PreviewType::Vector),
        "an enhanced metafile"
    );
    assert_eq!(
        classified("film.docx", &[0xD7, 0xCD, 0xC6, 0x9A, 0x00, 0x00]),
        Content::Kind(PreviewType::Vector),
        "a placeable Windows metafile"
    );
    assert_eq!(
        classified("film.docx", &[0x01, 0x00, 0x09, 0x00, 0x00, 0x01]),
        Content::Kind(PreviewType::Vector),
        "and a bare one"
    );

    // --------------------------------------------------------------- documents
    assert_eq!(
        classified("film.docx", b"@CT "),
        Content::Kind(PreviewType::Libre),
        "a Text602 document"
    );
    assert_eq!(
        classified("film.docx", b"RIFF\x00\x00\x00\x00CDR6"),
        Content::Kind(PreviewType::Libre),
        "a CorelDRAW drawing, in the RIFF container of the older versions"
    );
    assert_eq!(
        classified("film.docx", b"RIFF\x00\x00\x00\x00CMX1"),
        Content::Kind(PreviewType::Libre),
        "a Corel presentation exchange"
    );
    assert_eq!(
        classified("film.docx", b"BEGMF"),
        Content::Kind(PreviewType::Libre),
        "a Computer Graphics Metafile in its clear-text encoding"
    );
    assert_eq!(
        classified("film.docx", &[0x00, 0x20, 0x00, 0x00]),
        Content::Kind(PreviewType::Libre),
        "and one in the binary encoding"
    );
    assert_eq!(
        classified("film.docx", b"\x05\x07\x00\x00BOBO"),
        Content::Kind(PreviewType::Libre),
        "a ClarisWorks document"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 32];
            probe[0] = 0x03;
            probe[2] = 0x01;
            probe[3] = 0x0F;
            probe[8] = 0x41;
            probe
        }),
        Content::Kind(PreviewType::Libre),
        "a dBASE table, which is read for its version byte and its date"
    );
    assert_eq!(
        classified("film.docx", b"0\r\nSECTION\r\n"),
        Content::Kind(PreviewType::Libre),
        "a DXF drawing written as text"
    );
    assert_eq!(
        classified("film.docx", b"AutoCAD Binary DXF\r\n\x1a\x00"),
        Content::Kind(PreviewType::Libre),
        "and one written as binary"
    );
    assert_eq!(
        classified("film.docx", b"HWP Document File"),
        Content::Kind(PreviewType::Libre),
        "a Hangul document of the version that writes its name"
    );
    assert_eq!(
        classified("film.docx", b"WordPro"),
        Content::Kind(PreviewType::Libre),
        "a Lotus Word Pro document"
    );
    assert_eq!(
        classified("film.docx", &padded(2, &[0xD3, 0xA8, 0xA8])),
        Content::Kind(PreviewType::Libre),
        "an OS/2 metafile"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = padded(110, &[0x00, 0x06]);
            probe[6] = 0xFF;
            probe[7] = 0x99;
            probe
        }),
        Content::Kind(PreviewType::Libre),
        "a PageMaker 6 document"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = padded(110, &[0x32, 0x06]);
            probe[6] = 0xFF;
            probe[7] = 0x99;
            probe
        }),
        Content::Kind(PreviewType::Libre),
        "and a PageMaker 6.5 one, which the version word tells from it"
    );
    assert_eq!(
        classified("film.docx", b"VCLMTF"),
        Content::Kind(PreviewType::Libre),
        "a StarView metafile"
    );
    assert_eq!(
        classified("film.docx", b"ID;P"),
        Content::Kind(PreviewType::Libre),
        "a SYLK spreadsheet"
    );
    assert_eq!(
        classified("film.docx", b"\xFFWPC\x00\x00\x00\x00\x01\x0A"),
        Content::Kind(PreviewType::Libre),
        "a WordPerfect document"
    );
    assert_eq!(
        classified("film.docx", b"\xFFWPC\x00\x00\x00\x00\x01\x16"),
        Content::Kind(PreviewType::Libre),
        "and a WordPerfect drawing, which the file-type byte tells from it"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 22];
            probe[..2].copy_from_slice(&[0x01, 0xFE]);
            probe[20..22].copy_from_slice(&[0xD0, 0x02]);
            probe
        }),
        Content::Kind(PreviewType::Libre),
        "a Microsoft Works document"
    );
    assert_eq!(
        classified("film.docx", b"\x31\xBE\x00\x00"),
        Content::Kind(PreviewType::Libre),
        "a Windows Write document"
    );
    assert_eq!(
        classified("film.docx", b"\xFE\x37\x00\x1C\x00\x00\x00\x00"),
        Content::Kind(PreviewType::Libre),
        "a Word for the Macintosh document"
    );
    assert_eq!(
        classified("film.docx", b"\x09\x08\x08\x00\x00\x05"),
        Content::Kind(PreviewType::Libre),
        "a flat BIFF workbook"
    );
    for (name, head) in [
        ("123", b"\x00\x00\x1A\x00\x03\x10".as_slice()),
        ("wk3", b"\x00\x00\x1A\x00\x00\x10"),
        ("wk4", b"\x00\x00\x1A\x00\x02\x10"),
        ("wk1", b"\x00\x00\x02\x00\x06\x04"),
        ("wks", b"\x00\x00\x02\x00\x04\x04"),
        ("wb2", b"\x00\x00\x02\x00\x02\x10"),
        ("wq1", b"\x00\x00\x02\x00\x20\x51"),
        ("wq2", b"\x00\x00\x02\x00\x21\x51"),
    ] {
        assert_eq!(
            classified("film.docx", head),
            Content::Kind(PreviewType::Libre),
            "a `.{name}` spreadsheet of the years before the current ones"
        );
    }

    // ------------------------------------------------------------------ ebooks
    assert_eq!(
        classified("book.dat", &mobipocket()),
        Content::Kind(PreviewType::Calibre),
        "a Mobipocket book, which is what every Kindle format is inside"
    );
    assert_eq!(
        classified("book.dat", &declared_package("application/epub+zip")),
        Content::Kind(PreviewType::Calibre),
        "an EPub, named by the type it declares"
    );
    assert_eq!(
        classified("book.dat", b"<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<FictionBook xmlns=\"http://www.gribuser.ru/xml/fictionbook/2.0\">"),
        Content::Kind(PreviewType::Calibre),
        "a FictionBook, whose root element is what tells it from an ordinary XML file"
    );
    assert_eq!(
        classified(
            "book.dat",
            b"\xEF\xBB\xBF<?xml version=\"1.0\"?><FictionBook>"
        ),
        Content::Kind(PreviewType::Calibre),
        "and one written with a byte-order mark, which the declaration follows"
    );
    assert_eq!(
        classified("book.dat", b"AT&TFORMDJVM\x00\x00\x00\x08"),
        Content::Kind(PreviewType::Calibre),
        "a DjVu document"
    );
    assert_eq!(
        classified("book.dat", b"L\x00R\x00F\x00\x00\x00"),
        Content::Kind(PreviewType::Calibre),
        "a BBeB book, which writes its own letters with a zero byte between them"
    );

    // ------------------------------------------------------------------ videos
    assert_eq!(
        classified("film.docx", b"\x00\x00\x00\x20ftypisom"),
        Content::Kind(PreviewType::Videos),
        "an ISO base media file"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 188 * 4];
            for packet in 0..4 {
                probe[packet * 188] = 0x47;
            }
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a transport stream"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x01\xBA"),
        Content::Kind(PreviewType::Videos),
        "an MPEG program stream"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x01\xB3"),
        Content::Kind(PreviewType::Videos),
        "an MPEG elementary stream"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x00\x01\x67\x42\x00\x1E"),
        Content::Kind(PreviewType::Videos),
        "a raw H.264 stream"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x00\x01\x40\x01\x0C\x01\xFF\xFF"),
        Content::Kind(PreviewType::Videos),
        "a raw H.265 stream, whose video parameter set carries the field reserved to ones"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x00\x01\x00\x79"),
        Content::Kind(PreviewType::Videos),
        "a raw H.266 stream"
    );
    assert_eq!(
        classified("film.docx", b"\x12\x00\x0A\x02\x01\x02"),
        Content::Kind(PreviewType::Videos),
        "an AV1 stream, whose units have to chain"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x01\xB0\x20\x10\x0F\x00\x21\xC0"),
        Content::Kind(PreviewType::Videos),
        "an AVS sequence header, which is told from an MPEG-4 one by its profile"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x01\x0F\xC0"),
        Content::Kind(PreviewType::Videos),
        "a VC-1 sequence header of the advanced profile"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x01\x0F\x80"),
        Content::Unknown,
        "and one of a profile that carries no header of its own is left to the name"
    );
    assert_eq!(
        classified(
            "film.docx",
            b"BBCD\x00\x00\x00\x00\x0D\x00\x00\x00\x00BBCD\x00\x00\x00\x00\x00\x00\x00\x00\x0D"
        ),
        Content::Kind(PreviewType::Videos),
        "a Dirac or VC-2 stream, whose two units agree about where the first ended"
    );
    assert_eq!(
        classified("film.docx", b"aPv1"),
        Content::Kind(PreviewType::Videos),
        "an APV frame"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = Vec::new();
            for _ in 0..3 {
                probe.extend_from_slice(&[0x00, 0x00, 0x00, 0x02, 0x32, 0x01]);
            }
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "an EVC stream, whose units are length-prefixed"
    );
    assert_eq!(
        classified("film.docx", &padded(80, &[0x1F, 0x07, 0x00, 0x3F])),
        Content::Kind(PreviewType::Videos),
        "a DV stream"
    );
    assert_eq!(
        classified("film.docx", b"DKIF"),
        Content::Kind(PreviewType::Videos),
        "an IVF stream"
    );
    assert_eq!(
        classified(
            "film.docx",
            b"\x06\x0E\x2B\x34\x02\x05\x01\x01\x0D\x01\x02\x01\x01\x02"
        ),
        Content::Kind(PreviewType::Videos),
        "an MXF, and the `.imx` essence that is one"
    );
    assert_eq!(
        classified("film.docx", b"NUT/MULTI"),
        Content::Kind(PreviewType::Videos),
        "a NUT stream"
    );
    assert_eq!(
        classified("film.docx", b"YUV4MPEG2 W2 H2 F25:1 Ip A0:0 C420\n"),
        Content::Kind(PreviewType::Videos),
        "a YUV4MPEG2 stream"
    );
    assert_eq!(
        classified("film.docx", b"BIK"),
        Content::Kind(PreviewType::Videos),
        "a Bink movie"
    );
    assert_eq!(
        classified("film.docx", b"\x84\x10\xFF\xFF\xFF\xFF"),
        Content::Kind(PreviewType::Videos),
        "an Id RoQ movie"
    );
    assert_eq!(
        classified("film.docx", b"SMK2"),
        Content::Kind(PreviewType::Videos),
        "a Smacker movie"
    );
    assert_eq!(
        classified("film.docx", b"THP\x00"),
        Content::Kind(PreviewType::Videos),
        "a GameCube THP movie"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 36];
            probe[12..16].copy_from_slice(b"xobX");
            probe[16..20].copy_from_slice(&2u32.to_le_bytes());
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "an Xbox XMV movie"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 20];
            probe[..2].copy_from_slice(b"YO");
            probe[2] = 1;
            probe[3] = 2;
            probe[6] = 1;
            probe[7] = 1;
            probe[18..20].copy_from_slice(&920u16.to_le_bytes());
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a YOP movie, whose own fields are what its probe reads"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 20];
            probe[..4].copy_from_slice(b"RSD2");
            probe[8..12].copy_from_slice(&1u32.to_le_bytes());
            probe[16..20].copy_from_slice(&8000u32.to_le_bytes());
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a GameCube RSD stream"
    );
    assert_eq!(
        classified("film.docx", b".RMF\x00\x00"),
        Content::Kind(PreviewType::Videos),
        "a RealMedia stream"
    );
    assert_eq!(
        classified("film.docx", b".R1M\x00\x01\x01"),
        Content::Kind(PreviewType::Videos),
        "a recorded one"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 20];
            probe[..4].copy_from_slice(b"FILM");
            probe[16..20].copy_from_slice(b"FDSC");
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a Sega FILM movie"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 16];
            probe[4..6].copy_from_slice(&[0x01, 0xBC]);
            probe[10..16].copy_from_slice(&[0x00, 0x00, 0x00, 0x00, 0xE1, 0xE2]);
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a GXF stream"
    );
    assert_eq!(
        classified(
            "film.docx",
            &[
                0x11, 0xD2, 0xD3, 0xAB, 0xBA, 0xA9, 0xCF, 0x11, 0x8E, 0xE6, 0x00, 0xC0, 0x0C, 0x20,
                0x53, 0x65, 0x44
            ]
        ),
        Content::Kind(PreviewType::Videos),
        "an IFV stream"
    );
    assert_eq!(
        classified("film.docx", b"KDK\x00\x00"),
        Content::Kind(PreviewType::Videos),
        "a KUX movie, which is a Flash movie with a header of its own"
    );
    assert_eq!(
        classified("film.docx", b"Interplay MVE File\x1A\x00\x1A\x00"),
        Content::Kind(PreviewType::Videos),
        "an Interplay MVE movie"
    );
    assert_eq!(
        classified("film.docx", b"pmpm\x01\x00\x00\x00"),
        Content::Kind(PreviewType::Videos),
        "a PMP movie"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 36];
            probe[..4].copy_from_slice(b"CRID");
            probe[32..36].copy_from_slice(b"@UTF");
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a Scaleform movie, which FFmpeg has no demuxer for and `file` names"
    );
    assert_eq!(
        classified("film.docx", b"DAHUA"),
        Content::Kind(PreviewType::Videos),
        "a Dahua camera's stream"
    );
    assert_eq!(
        classified("film.docx", b"\x00abcVersion:Vivo/0"),
        Content::Kind(PreviewType::Videos),
        "a Vivo stream"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 28];
            probe[3] = 0xC5;
            probe[4..8].copy_from_slice(&8u32.to_le_bytes());
            probe[24..28].copy_from_slice(&[0x0C, 0x00, 0x00, 0x00]);
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a VC-1 test stream"
    );
    assert_eq!(
        classified(
            "film.docx",
            &[0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00]
        ),
        Content::Kind(PreviewType::Videos),
        "a PlayStation STR stream"
    );
    assert_eq!(
        classified(
            "film.docx",
            &[
                0xB7, 0xD8, 0x00, 0x20, 0x37, 0x49, 0xDA, 0x11, 0xA6, 0x4E, 0x00, 0x07, 0xE9, 0x5E,
                0xAD, 0x8D
            ]
        ),
        Content::Kind(PreviewType::Videos),
        "a Windows recorded television stream"
    );
    assert_eq!(
        classified("film.docx", b"NSVf"),
        Content::Kind(PreviewType::Videos),
        "a Nullsoft stream"
    );

    // The formats whose probe scores a shape rather than a magic, each with the shape
    // FFmpeg's own demuxer scores: a block table, a periodic command byte, a header of
    // offsets, a table that has to chain, a picture start code twice over.
    assert_eq!(
        classified(
            "film.docx",
            &[
                0x01, 0x00, 0x04, 0x01, 0x05, 0x00, 0x02, 0x01, 0x07, 0x00, 0x03, 0x01, 0x0A, 0x00,
                0x01, 0x01
            ]
        ),
        Content::Kind(PreviewType::Videos),
        "an Interplay C93 movie"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 24 * 8];
            for packet in 0..8 {
                probe[packet * 24] = 0x09;
            }
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a CD Graphics stream"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 32];
            probe[2..6].copy_from_slice(&1388u32.to_be_bytes());
            probe[14..16].copy_from_slice(&320u16.to_be_bytes());
            probe[16..18].copy_from_slice(&200u16.to_be_bytes());
            probe[19] = 6;
            probe[20..22].copy_from_slice(&256u16.to_be_bytes());
            probe[22..24].copy_from_slice(&100u16.to_be_bytes());
            probe[24..26].copy_from_slice(&11025u16.to_be_bytes());
            probe[26] = 15;
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a Commodore CDXL stream"
    );
    assert_eq!(
        classified("film.docx", &{
            let mut probe = vec![0u8; 26];
            probe[..2].copy_from_slice(b"L2");
            probe[12..14].copy_from_slice(&1u16.to_be_bytes());
            probe[14..18].copy_from_slice(&[0x00, 0x01, 0x00, 0x04]);
            probe
        }),
        Content::Kind(PreviewType::Videos),
        "a Moflex movie"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x01\x00\x11\x22\x00\x01\x00"),
        Content::Kind(PreviewType::Videos),
        "an H.261 stream"
    );
    assert_eq!(
        classified("film.docx", b"\x00\x00\x80\x11\x00\x00\x80"),
        Content::Kind(PreviewType::Videos),
        "an H.263 stream"
    );

    // -------------------------------------------------------------------- text
    assert_eq!(
        classified("film.docx", b"{\\rtf1\\ansi"),
        Content::Kind(PreviewType::Text),
        "an RTF document, which the text preview shows"
    );
    assert_eq!(
        classified("film.docx", b"ttcf\x00\x01\x00\x00"),
        Content::Kind(PreviewType::Fonts),
        "a collection of fonts"
    );
    assert_eq!(
        classified("film.docx", b"%!PS-Adobe-3.0"),
        Content::Kind(PreviewType::Vector),
        "and a PostScript drawing"
    );
}
