//! Which files are previewed by ImageMagick rather than by a reader of this app's own.
//!
//! ImageMagick is the one tool that opens nearly every picture ever written: its own coders for
//! the formats nobody else kept, and the camera raw formats through the raw decoder it is built
//! with — LibRaw, the library behind `dcraw`. A camera writes a `.nef`, a `.cr3`, a `.arw`, a
//! `.raf` or a `.dng` — a raw sensor reading with a picture wrapped around it, which no decoder
//! of this app reads and Windows has no codec for — and where ImageMagick is installed this app
//! asks it to draw one as a picture (see `imagemagick_render`). What is listed here is which of
//! those names this app hands over.
//!
//! The list is built from the engine's own registry rather than from what it is said to support:
//! every name in it was read one at a time out of `magick -list format` on an installed engine,
//! and a name belongs here only where that registry declares read support for it. A name the
//! engine has no coder for is a launch that answers nothing, and a name it reads through a
//! delegate the machine has not been given is the same launch — what that costs is bounded by the
//! engine remembering the answer (see `imagemagick_render`).
//!
//! What is *not* here is a judgement about the engine or the format rather than a gap, and each
//! group is worth naming:
//!
//! * **A name another kind already reads is not here.** `pcx`, `pcd`, `pct`, `ras`, `wpg`, `sti`
//!   and `cin` are entries of the `[libre]` and video lists, `psd`, `exr`, `hdr`, `dds`, `qoi`
//!   and `svg` of the readers this app carries, and a name sits in exactly one list so that a
//!   preview of one cannot come back by two routes. The *spellings* of a format are not that
//!   rule, which is why several names below sit beside a sibling in another list: `pict` is here
//!   while `pct` is the `[libre]` list's, `sun` is here while `ras` is, and `pcds` is here while
//!   `pcd` is — the engine the other list names has no filter for the second spelling, so that
//!   spelling is a name no other list claims. `dxt1` and `dxt5` are the same case against the
//!   picture path: a `.dds` is decoded by a reader of this app's, and its two spellings are not.
//!   One name here is asked for and may draw nothing, which is the coder's own limit rather than
//!   a delegate's absence: measured against the installed build, `pict` reads a QuickDraw picture
//!   of a few dozen pixels and refuses one of a few hundred with `insufficient image data`, so a
//!   `.pict` of a size anyone would hover is a hover with no preview rather than a broken one.
//! * **The fonts are not here.** `pfa`, `pfb` and `dfont` are fonts the browser engine the
//!   specimens are drawn by cannot be given — every font a page asks for goes through that
//!   engine's own sanitizer, which takes OpenType and the two WebFont containers and nothing else
//!   (see `font_formats`, where the measurement is) — and the answer to that is not this engine:
//!   a font is the font preview's kind, and a specimen is a page of text rather than a picture,
//!   so a rendering of a face by an image converter is not what this app wants for one.
//! * **The documents are not here.** A PostScript or PDF file, an `.xps`, a `.djvu` and an
//!   `.mvg` are pages rather than pictures, and the engine draws them only through a Ghostscript,
//!   a DjVuLibre or a Graphviz installed beside it. A page of a document is what the PDF path and
//!   the render engine are for, and a name here answers nothing on a machine that has none of
//!   those delegates — which is every machine this was measured on. A user who has one can add
//!   the name by hand.
//! * **The things that are not files are not here.** `xc`, `canvas`, `caption`, `gradient`,
//!   `label`, `null`, `pattern`, `plasma`, `tile`, `http`, `https`, `ftp`, `file`, `inline`,
//!   `data`, `clipboard`, `vid`, `screenshot`, `thumbnail`, `mask`, `clip`, `msvg`, `rsvg` and
//!   `dcraw` are the engine's own notation: names of images to *make* rather than of files to
//!   open, and a hover on a file called `xc` is a hover on nothing. What the text lists answer for
//!   — `txt`, `html`, `json`, `yaml` — is theirs, and `info`, `kernel`, `histogram`, `matte`,
//!   `uil`, `shtml`, `brf`, `cip`, `isobrl`, `isobrl6`, `ubrl`, `ubrl6`, `ashlar`, `eps2`, `eps3`,
//!   `ps2` and `ps3` are shapes the engine only ever *writes*: a name behind one of those is a
//!   file nothing here can open, and one the engine itself cannot read.
//! * **And one name is missing because the build is.** `xwd` is an X11 window dump, a format
//!   ImageMagick has a coder for (registered as `_XWD`) that the Windows installer's module set
//!   does not ship: asking one of those builds for a `.xwd` fails looking for a coder module that
//!   is not there, and `magick -list format` — which names the formats whose coders are installed
//!   — does not name it at all. A build that ships it declares it, and then the name is an entry
//!   away like every other.
//!
//! A raw sample dump is the one kind of file here the engine cannot measure for itself. `rgb`,
//! `rgba`, `gray`, `cmyk` and the rest of them are samples with no header at all — no size, no
//! depth, no channel order beyond the name — because the shape of one was written down where the
//! file was made rather than in the file. What is left to read is the file's own length, so the
//! shape is worked out from that and from what one pixel of the name weighs, and a length that
//! does not settle it is answered with no preview rather than with a guess (see
//! `imagemagick_render::raw_geometry`); the engine is then told the size the file is to be read
//! at.
//!
//! The list of names this answers for is a row of `crate::formats::lists` — the one table
//! every kind's list is a row of, and the one place a list is written down, the built-in
//! entries and the older lists this app shipped and then changed included.

use crate::config::config::PreviewType;
use crate::formats::lists;
use crate::CONFIG;
use std::path::Path;

/// Whether the engine is the one that develops this file at all: a name its own list carries,
/// or the bytes of a picture it reads under a name no list holds.
///
/// It is the question the request side asks before it asks the engine for a picture — a
/// drawing renamed to `.bin` is still a picture the engine reads, and one whose bytes name
/// another kind is not one to start it for — and it is the same question asked the same way
/// `libre_formats::engine_page_kind` asks it for the render engine: the file's own bytes
/// first, the name it carries after them. What it is *not* is a gate: whether that kind may be
/// shown is the caller's to ask (`PreviewType::enabled`), because the same question is asked
/// of a preview that is already on screen when a switch is thrown.
pub fn is_engine_picture(path: &Path) -> bool {
    use crate::formats::content_type::Content;

    // The entry is read before the lock, so the guard is not held across the read (see
    // `content_type::of_reaching_config`), and this list comparison is in memory under a guard
    // of its own.
    let content = crate::formats::content_type::of_reaching_config(path);

    match content {
        Content::Kind(PreviewType::Magick) => true,
        Content::Kind(_) | Content::Foreign => false,
        Content::Unknown => CONFIG
            .lock()
            .map(|config| lists::MAGICK.claims(path, &config))
            .unwrap_or(false),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// What the list holds is what this app has no reader for: a picture, a drawing, a
    /// font, a PDF and a text file are all answered elsewhere, and none of them is here.
    /// The names it shares a format with but not a spelling are the reason the rule is
    /// worth stating rather than assuming.
    #[test]
    fn holds_no_name_another_kind_already_reads() {
        let config = crate::config::config::AppConfig::default();

        for name in [
            "photo.png",
            "photo.jpg",
            "photo.heic",
            "photo.avif",
            "photo.jxl",
            "render.exr",
            "light.hdr",
            "texture.dds",
            "still.qoi",
            "shot.tga",
            "scan.tif",
            "icon.ico",
            "drawing.svg",
            "drawing.emf",
            "print.eps",
            "report.pdf",
            "notes.txt",
            "readme.md",
            "page.html",
            "font.ttf",
            "font.otf",
            "bundle.zip",
            "animation.swf",
            "clip.mp4",
            "drawing.cdr",
            "painting.psd",
            "project.kra",
            "scan.pcx",
            "scan.pcd",
            "drawing.pct",
            "image.ras",
            "drawing.wpg",
            "sheet.sti",
            "frame.cin",
            "picture.palm",
        ] {
            let path = std::path::Path::new(name);
            let claimed = crate::formats::lists::IMAGE.claims(path, &config)
                || crate::formats::lists::VECTOR.claims(path, &config)
                || crate::formats::lists::DESIGN.claims(path, &config)
                || crate::formats::lists::FONT.claims(path, &config)
                || crate::formats::lists::LIBRE.claims(path, &config)
                || crate::formats::video_formats::matches_any_video_list(path, &config)
                || crate::formats::lists::ARCHIVE.claims(path, &config)
                || crate::formats::lists::OFFICE.claims(path, &config)
                || crate::formats::lists::TEXT.claims(path, &config)
                || crate::formats::lists::NAMES.claims(path, &config);

            if !claimed {
                continue;
            }

            assert!(
                !crate::formats::lists::MAGICK.claims(path, &config),
                "`{name}` is read by another kind, so the engine is not asked about it"
            );
        }
    }

    /// And the names it does hold are the camera raws above all, which is what the list
    /// exists for: nothing else on the machine opens one.
    #[test]
    fn holds_the_camera_raw_formats_and_the_pictures_beside_them() {
        let config = crate::config::config::AppConfig::default();

        for name in [
            "shot.nef",
            "shot.NEF",
            "shot.cr2",
            "shot.cr3",
            "shot.arw",
            "shot.dng",
            "shot.raf",
            "shot.orf",
            "shot.rw2",
            "shot.pef",
            "shot.srw",
            "shot.x3f",
            "shot.3fr",
            "shot.raw",
            "scan.dcm",
            "frame.dpx",
            "sky.fits",
            "sky.fit",
            "sky.fts",
            "still.jp2",
            "still.j2k",
            "still.jng",
            "still.mng",
            "picture.xcf",
            "picture.sgi",
            "picture.xbm",
            "picture.xpm",
            "picture.miff",
            "picture.pfm",
            "picture.vicar",
            "picture.wbmp",
            "cursor.cur",
            "fax.dcx",
        ] {
            assert!(
                crate::formats::lists::MAGICK.claims(std::path::Path::new(name), &config),
                "`{name}` is one of the engine's pictures"
            );
        }
    }

    /// The list is what the user edits, and an entry that is not a bare extension is
    /// dropped rather than matched against.
    #[test]
    fn reads_a_list_of_bare_extensions() {
        let extensions =
            crate::formats::lists::sanitize_extension_list(" .NEF , cr2,,*.raw ,nef,3fr");

        assert_eq!(extensions, vec!["nef", "cr2", "3fr"]);
    }

    /// What the engine is asked about is asked of the file's own bytes before its name, which
    /// is what makes a picture renamed to a name no list holds the engine's to develop: a
    /// Silicon Graphics picture under a `.bin` is one of its pictures, a picture under a
    /// `.png` is a picture and not the engine's, and a name nothing recognizes is left to the
    /// list the name is written in.
    #[test]
    fn asks_the_engine_about_a_file_by_its_bytes_before_its_name() {
        let folder = std::env::temp_dir().join("rust-hover-preview-magick-engine-picture");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // A Silicon Graphics picture, which is `01 DA` and the storage and depth bytes every
        // one of them carries.
        let renamed = folder.join("artwork.bin");
        std::fs::write(&renamed, [0x01, 0xDA, 0x00, 0x02, 0x00, 0x03, 0x06, 0x40])
            .expect("a written picture");
        assert!(
            is_engine_picture(&renamed),
            "a picture the engine reads is its to develop, whatever it is called"
        );

        // The same bytes under a picture's own name are a picture of this app's: the two
        // agree, so there is nothing for the engine to be asked about.
        let named = folder.join("artwork.png");
        std::fs::write(&named, [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
            .expect("a written picture");
        assert!(
            !is_engine_picture(&named),
            "a picture this app decodes itself is not one to start an engine for"
        );

        // And the name's own list, for the files whose bytes say nothing at all.
        std::fs::write(folder.join("shot.nef"), b"a raw, of a sort").expect("a written file");
        assert!(
            is_engine_picture(&folder.join("shot.nef")),
            "a name the list carries is the engine's, whatever the bytes say"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }
}
