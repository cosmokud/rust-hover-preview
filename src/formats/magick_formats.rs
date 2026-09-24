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
//! The list lives in `config.ini` as `[magick] extensions`, written from the built-in list
//! on first run and read back from there, so a user can add a format the engine reads and
//! this app does not know, or take one out. The question here is only what a file is
//! *called*: whether the engine can read it at all is settled by the engine, and a name it
//! cannot read is answered with no preview — once, and then remembered, so a name that was
//! put in this list by mistake costs one conversion and never another.

use crate::config::config::PreviewType;
use crate::formats::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the picture formats ImageMagick
/// reads and this app has no reader of its own for.
///
/// The camera raw formats come first in spirit if not in order — `3fr`, `arw`, `cr2`, `cr3`,
/// `crw`, `dcr`, `dng`, `erf`, `fff`, `iiq`, `k25`, `kdc`, `mdc`, `mef`, `mos`, `mrw`,
/// `nef`, `nrw`, `orf`, `pef`, `raf`, `raw`, `rmf`, `rw2`, `rwl`, `sr2`, `srf`, `srw` and
/// `x3f`, from the raw decoder the engine is built with, which is every raw format a camera
/// writes that is still met with — and they are what this list is mostly about: nothing else
/// on a Windows machine opens one, the shell shows a thumbnail from the picture the camera
/// left inside the file and nothing else, and a hover onto one shows nothing at all rather
/// than the photograph.
///
/// What follows them, in the order the list above is written in:
///
/// * The formats the engine has a coder of its own for and no other list claims: the medical
///   scanner's `dcm`, the film scanner's `dcx` and `dpx`, the astronomer's `fit`, `fits` and
///   `fts`, the compositor's `j2c`, `j2k`, `jp2`, `jpc`, `jpm` and `jpt`, the animator's `jng`
///   and `mng`, the illustrator's `xcf`, `xbm` and `xpm`, the painter's `sgi`, the engineer's
///   `vicar`, the phone's `wbmp`, the cursor's `cur`, and the engine's own `miff`.
/// * The formats the registry declares read support for that no preview was ever asked for: a
///   pixel-art program's sprites (`ase`, `aseprite`), a floppy-era paint program's pictures
///   (`mac`, `pix`, `rla`, `rle`, `art`, `cut`, `wbinfo`), a fax machine's pages (`fax`, `g3`,
///   `g4`), a scanner's `pgx`, a texture of a console's (`tim`, `tm2`), an icon of a robot's
///   (`rgf`), an embroidery machine's pattern (`pes`), a spectrum analyser's screen (`scr`), a
///   telescope's or a satellite's frame (`hrz`, `ipl`, `fl32`, `sct`, `jnx`), a colour lookup
///   table (`cube`), a document's provenance record (`c2pa`), a markup language (`pango`), and
///   the rest of the names beside them. Each is a picture the engine really draws; what none of
///   them has is a reason to have been picked over the others, and the list shipping with them
///   is what settles that.
/// * The second spelling of a format whose first spelling another list holds — `pict`, `sun`,
///   `pcds`, `dxt1` and `dxt5`, `icb`, `vda` and `vst`, `picon` — which the engine reads and the
///   engine the other spelling belongs to does not (see the module documentation above).
/// * The raw sample dumps (`rgb`, `rgba`, `gray`, `cmyk` and their kin, and the `group4` fax
///   bitstream), whose shape comes from their own length: a dump whose length does not settle one
///   is a file this app shows nothing for, and one that does is read at the size that came out of
///   it.
///
/// What is deliberately *not* here is named in the module documentation above: the alternatives
/// spelling of a name another list already claims, the fonts, the documents that need a delegate,
/// the engine's own notation, and `xwd`, which this build ships no coder for.
pub const DEFAULT_MAGICK_EXTENSIONS: &str = "3fr,aai,art,arw,ase,aseprite,bayer,bayera,bgr,bgra,bgro,c2pa,cal,cals,cmyk,cmyka,cr2,cr3,crw,cube,cur,cut,\
dcm,dcr,dcx,dng,dpx,dxt1,dxt5,erf,fax,fff,fit,fits,fl32,fts,ftxt,g3,g4,gray,graya,group4,hrz,icb,iiq,ipl,j2c,\
j2k,jng,jnx,jp2,jpc,jpm,jpt,k25,kdc,mac,map,mat,mdc,mef,miff,mng,mono,mos,mpc,mrw,mtv,nef,nrw,orf,otb,pal,\
palm,pango,pcds,pef,pes,pfm,pgx,phm,picon,pict,pix,pwp,raf,raw,rgb,rgb565,rgba,rgbo,rgf,rla,rle,rmf,rw2,rwl,\
scr,sct,sf3,sfw,sgi,six,sixel,sr2,srf,srw,stegano,sun,tim,tm2,uyvy,vda,vicar,viff,vips,vst,wbinfo,wbmp,x3f,\
xbm,xcf,xpm,xv,ycbcr,ycbcra,yuv";

/// The built-in list as it stood while it held the camera raw formats and the pictures beside them
/// and nothing else: what this app wrote before the names nobody had asked for were added to it,
/// and before the raw sample dumps were.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written,
/// which is what gives an installation that already exists the names added since. Without it those
/// names would reach a fresh installation only, since every `config.ini` already written holds the
/// list as it was (see `config::repair_older_lists`).
pub const MAGICK_EXTENSIONS_BEFORE_THE_REST: &str = "3fr,arw,cr2,cr3,crw,cur,dcm,dcr,dcx,dng,dpx,erf,fff,fit,fits,fts,iiq,j2c,j2k,jng,jp2,jpc,jpm,jpt,k25,kdc,mdc,mef,miff,mng,mos,mrw,nef,nrw,orf,pef,pfm,raf,raw,rmf,rw2,rwl,sgi,sr2,srf,srw,vicar,wbmp,x3f,xbm,xcf,xpm";

/// Whether the configured list claims `path`.
pub fn matches_magick_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups use.
///
/// Every name in this list is a bare extension — unlike the archive list, which has to
/// carry the dotted `tar.gz` — so anything that is not one is dropped rather than matched
/// against.
pub fn sanitize_magick_extensions(list: &str) -> Vec<String> {
    let mut extensions: Vec<String> = Vec::new();

    for entry in list.split(',') {
        let trimmed = entry.trim().trim_start_matches('.').to_lowercase();
        let is_extension = !trimmed.is_empty()
            && !trimmed.starts_with('.')
            && !trimmed.ends_with('.')
            && trimmed
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '-' | '_'));

        if is_extension && !extensions.contains(&trimmed) {
            extensions.push(trimmed);
        }
    }

    extensions
}

/// Whether the configured list claims `path`, without asking whether these previews are
/// switched on. The gate is asked beside it by the hook, the way every other kind's is.
pub fn is_magick_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_magick_list(path, &config.magick_extensions))
        .unwrap_or(false)
}

/// Whether a preview may be shown for `path`: the file the configured list claims, and
/// the `Magick` gate in the tray's `Preview Types` submenu.
///
/// Both halves ask it where a kind can be switched off under a preview that is already on
/// screen: a hover is not sent for a kind that is off, and the layout places nothing for a
/// file whose kind is off, which is how a preview of that kind comes down when the switch
/// does. See `PreviewType::enabled`.
pub fn is_magick_preview(path: &Path) -> bool {
    is_magick_file(path) && PreviewType::Magick.enabled()
}

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
    use crate::formats::content_type::{self, Content};

    match content_type::of(path) {
        Content::Kind(PreviewType::Magick) => true,
        Content::Kind(_) | Content::Foreign => false,
        Content::Unknown => is_magick_file(path),
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
        let list = sanitize_magick_extensions(DEFAULT_MAGICK_EXTENSIONS);
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
            let claimed = crate::formats::image_formats::matches_image_list(&path, &config.image_extensions)
                || crate::formats::vector_formats::matches_vector_list(&path, &config.vector_extensions)
                || crate::formats::design_formats::matches_design_list(&path, &config.design_extensions)
                || crate::formats::font_formats::matches_font_list(&path, &config.font_extensions)
                || crate::formats::libre_formats::matches_libre_list(&path, &config.libre_extensions)
                || crate::formats::video_formats::matches_video_list(&path, &config.video_extensions)
                || crate::formats::archive_formats::matches_archive_list(&path, &config.archive_extensions)
                || crate::formats::office_formats::matches_office_list(&path, &config.office_extensions)
                || crate::formats::text_formats::matches_text_lists(
                    &path,
                    &config.text_extensions,
                    &config.text_names,
                );

            if !claimed {
                continue;
            }

            assert!(
                !matches_magick_list(path, &list),
                "`{name}` is read by another kind, so the engine is not asked about it"
            );
        }
    }

    /// And the names it does hold are the camera raws above all, which is what the list
    /// exists for: nothing else on the machine opens one.
    #[test]
    fn holds_the_camera_raw_formats_and_the_pictures_beside_them() {
        let list = sanitize_magick_extensions(DEFAULT_MAGICK_EXTENSIONS);

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
                matches_magick_list(std::path::Path::new(name), &list),
                "`{name}` is one of the engine's pictures"
            );
        }
    }

    /// The list is what the user edits, and an entry that is not a bare extension is
    /// dropped rather than matched against.
    #[test]
    fn reads_a_list_of_bare_extensions() {
        let extensions = sanitize_magick_extensions(" .NEF , cr2,,*.raw ,nef,3fr");

        assert_eq!(extensions, vec!["nef", "cr2", "3fr"]);
    }

    /// What the engine is asked about is asked of the file's own bytes before its name, which
    /// is what makes a picture renamed to a name no list holds the engine's to develop: a
    /// Silicon Graphics picture under a `.bin` is one of its pictures, a picture under a
    /// `.png` is a picture and not the engine's, and a name nothing recognizes is left to the
    /// list the name is written in.
    #[test]
    fn asks_the_engine_about_a_file_by_its_bytes_before_its_name() {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.confirm_file_type = true;
        }

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
        std::fs::write(&named, [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]).expect("a written picture");
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
