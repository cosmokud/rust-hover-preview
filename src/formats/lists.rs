//! Every extension list this app holds, and the one place a list is written down.
//!
//! Sixteen lists — fifteen kinds and the name list the text preview reads a dot file by —
//! were sixteen files whose whole content was a `pub const` of comma-separated extensions and
//! three lines that read the list back and handed the answer to a gate. Adding a kind meant
//! writing that file, importing its default list into `config`, coercing its sanitiser to a
//! `fn(&str) -> Vec<String>` so an array of the sixteen would typecheck at all, taking the list
//! out of `AppConfig` and putting it back again for the two resets, and naming the field in the
//! read that fills it from the file: six edits across three files, and the four of them that
//! were not the list itself had nothing in them to fail when one was forgotten.
//!
//! Two of the sixteen had to be read by a rule of their own — the archive list holds a dotted
//! compound name, and the name list holds a file name rather than an extension — and each of
//! those rules was written out as its own sanitiser rather than as a parameter, because the
//! thirteen that were byte-identical were copied rather than shared, and a copy is a copy that
//! can drift from the twelve beside it.
//!
//! So the lists are rows here. A row says where the file writes the list, what it holds on the
//! first run, what this app held of it before, what an entry of it may be, and which field of
//! the configuration it is; everything that used to name those things one list at a time — the
//! read out of a file, the write into one, the two resets, the repair that brings a file an older
//! build wrote up to the list of now, and the lookup each list is asked through — walks this table
//! instead, and a kind added to the app is a row added to it.

use crate::config::config::AppConfig;
use configparser::ini::Ini;
use std::path::Path;

/// What the entries of one list are allowed to be, which is the whole of what sixteen
/// sanitizers differed on once fourteen of them were one function.
///
/// A choice of characters rather than of function on purpose: a row that named a sanitiser could
/// name one that does not exist, and the difference between the four was never a different
/// function — it was a different set of characters to read an entry by.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Entries {
    /// A bare extension: its alphanumerics and `+`, `-` and `_`. Twelve of the sixteen lists are
    /// this and nothing else, and every one of them reads the same way.
    Bare,
    /// A bare extension, or a compound of two of them. The archive list alone, because `tar.gz`
    /// is a name rather than an extension — the last dot of `sources.tar.gz` is `gz`, which is
    /// not an archive on its own — so a sanitiser that dropped the dot would silently stop the
    /// list claiming the format it exists to claim.
    Compound,
    /// A bare extension, or one holding a `#`. The text list alone, and only so that a C# or F#
    /// project is a text file rather than an extension nobody can type.
    Hashed,
    /// A file name rather than an extension: a dot begins the name instead of changing it, so
    /// `.gitignore` has no extension at all and `cmakelists.txt` has one in the middle of it.
    Name,
}

impl Entries {
    /// One list as the entries a lookup compares against, by the rule the row names.
    ///
    /// Every list in the running app is read through here and nowhere else — the four rules
    /// below are the four this dispatches to, under the names the rest of the tree has always
    /// called them by.
    fn sanitize(self, list: &str) -> Vec<String> {
        match self {
            Entries::Bare => sanitize_extension_list(list),
            Entries::Compound => sanitize_archive_extension_list(list),
            Entries::Hashed => sanitize_extensions(list),
            Entries::Name => sanitize_names(list),
        }
    }
}

/// The characters an ordinary extension list admits besides its alphanumerics.
///
/// Deliberately without `#` and without `.`, which are the two rules that need them: `#` for a
/// name like `C#`, `.` for a compound like `tar.gz`.
const PLAIN_EXTENSION_CHARS: &[char] = &['+', '-', '_'];
/// The archive list's own set, which is the one list that may hold a dotted compound name.
const COMPOUND_EXTENSION_CHARS: &[char] = &['.', '+', '-', '_', '#'];
/// The text list's own set, which is the one extension list that may hold a `#`.
const HASHED_EXTENSION_CHARS: &[char] = &['+', '-', '_', '#'];
/// And the name list's, for entries that are file names rather than extensions.
const NAME_CHARS: &[char] = &['.', '-', '_'];

/// One list split into the entries a lookup compares against, each read by the rule `allowed`
/// names.
///
/// A leading dot is accepted because `py` and `.py` are both what a user might type, and an entry
/// that is not a plausible name or extension is dropped so a stray path or sentence in the list
/// cannot turn into a match — which is what an entry is read as everywhere else in this app.
fn sanitized_list(list: &str, allowed: &[char]) -> Vec<String> {
    let mut entries: Vec<String> = Vec::new();

    for entry in list.split(',') {
        let entry = entry.trim().trim_start_matches('.').to_lowercase();
        if entry.is_empty()
            || !entry
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || allowed.contains(&c))
        {
            continue;
        }

        // A list is a set, and a user who typed a name twice gets it once: the second entry would
        // be a second match where one file already satisfies the first.
        if !entries.iter().any(|held| held == &entry) {
            entries.push(entry);
        }
    }

    entries
}

/// The sanitiser for the twelve lists that are a bare extension and nothing else.
pub fn sanitize_extension_list(list: &str) -> Vec<String> {
    sanitized_list(list, PLAIN_EXTENSION_CHARS)
}

/// The archive list's own sanitiser, which is the one list that holds a dotted compound name.
///
/// `tar.gz` is a name rather than an extension, and `archive_formats::matches_archive_list`
/// matches it against the end of a whole file name — so a sanitiser that dropped the dot would
/// silently stop the list claiming the format it exists to claim.
pub fn sanitize_archive_extension_list(list: &str) -> Vec<String> {
    sanitized_list(list, COMPOUND_EXTENSION_CHARS)
}

/// The text list's own sanitiser, which is the one extension list that admits a `#`, for `C#`.
pub fn sanitize_extensions(list: &str) -> Vec<String> {
    sanitized_list(list, HASHED_EXTENSION_CHARS)
}

/// The name list's own sanitiser, for a list whose entries are file names and not extensions.
pub fn sanitize_names(list: &str) -> Vec<String> {
    sanitized_list(list, NAME_CHARS)
}

/// The extensions written to `config.ini` on first run: the still and animated picture formats a
/// hover is expected to meet.
///
/// `apng` is a name like any other here — an animated PNG is recognized by its own `acTL` chunk
/// rather than by what it is called — and `gif`, `png` and `webp` are each both an animated
/// format and a still one. `avci`, `avif`, `heic`, `heif`, `jxl` and a still `webp` are pictures
/// the same way, and what decodes them is the codec Windows has rather than one this app carries
/// — `avci` being the AVC still of the same container family, beside the HEVC and AV1 ones — a
/// `webp` being the one of them with a second reader behind it, libwebp in the binary, for the
/// machine that codec is missing from and for the picture that moves; see `wic_image` and
/// `webp_image`. `dds` is a picture of the same kind and by the same route: the codec Windows has
/// for it draws one at the size of the preview rather than the size of the file, and the formats
/// that codec does not read — the uncompressed ones, BC4 and BC5 — are read by a decoder of this
/// app's own; see `dds_image`.
///
/// `svg` and `svgz` are not entries of this list, though they were while the kind they are under
/// was called SVG: what draws one is a browser rather than a decoder, so they are entries of the
/// vector list and are gated — and sized, and drawn over — by that kind; see `vector_formats` and
/// `svg_preview`.
pub const DEFAULT_IMAGE_EXTENSIONS: &str =
    "apng,avci,avif,bmp,dds,exr,ff,gif,hdr,heic,heif,ico,jfif,jpe,jpeg,jpg,jxl,pam,pbm,pgm,png,pnm,ppm,qoi,tga,tif,tiff,webp";

/// The built-in image list as it stood while `svg` and `svgz` were entries of it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written,
/// and the two entries it names are left to the vector list, which is where a document belongs.
/// A list anyone has edited is kept as it is, and an `svg` named by one is still drawn by the
/// browser and still gated by the `Vector` kind, because what a file is, is its own answer
/// rather than the list's (see `explorer_hook::is_media_file`).
pub const IMAGE_EXTENSIONS_WITH_SVG: &str =
    "apng,avif,bmp,dds,exr,ff,gif,hdr,heic,heif,ico,jfif,jpe,jpeg,jpg,jxl,pam,pbm,pgm,png,pnm,ppm,qoi,svg,svgz,tga,tif,tiff,webp";

/// The built-in image list as it stood before `dds` was added to it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the format would reach a fresh installation only: every `config.ini` already
/// written holds a list that differs from the built-in one, and a list that differs is otherwise
/// the user's own (see `repair_older_lists`).
pub const IMAGE_EXTENSIONS_BEFORE_DDS: &str =
    "apng,avif,bmp,exr,ff,gif,hdr,heic,heif,ico,jfif,jpe,jpeg,jpg,jxl,pam,pbm,pgm,png,pnm,ppm,qoi,svg,svgz,tga,tif,tiff,webp";

/// The built-in image list as it stood before `avci` was added to it — the AVC still of the
/// container family whose HEVC and AV1 stills the codec Windows has reads under `heic` and
/// `avif`.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the name would reach a fresh installation only: every `config.ini` already
/// written holds the list as it was, and a list that differs is otherwise the user's own (see
/// `repair_older_lists`).
pub const IMAGE_EXTENSIONS_BEFORE_AVCI: &str =
    "apng,avif,bmp,dds,exr,ff,gif,hdr,heic,heif,ico,jfif,jpe,jpeg,jpg,jxl,pam,pbm,pgm,png,pnm,ppm,qoi,tga,tif,tiff,webp";

/// The built-in image list as it stood before the formats Windows has a codec for were added to
/// it: `avif`, `heic`, `heif` and `jxl`.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the four formats would reach a fresh installation only: every `config.ini`
/// already written holds a list that differs from the built-in one, and a list that differs is
/// otherwise the user's own (see `repair_older_lists`).
pub const IMAGE_EXTENSIONS_BEFORE_CODEC_FORMATS: &str =
    "apng,bmp,exr,ff,gif,hdr,ico,jfif,jpe,jpeg,jpg,pam,pbm,pgm,png,pnm,ppm,qoi,svg,svgz,tga,tif,tiff,webp";

/// The built-in image list as it stood before `svg` and `svgz` were added to it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the two formats would reach a fresh installation only: every `config.ini`
/// already written holds a list that differs from the built-in one, and a list that differs is
/// otherwise the user's own (see `repair_older_lists`).
pub const IMAGE_EXTENSIONS_BEFORE_SVG: &str =
    "apng,bmp,exr,ff,gif,hdr,ico,jfif,jpe,jpeg,jpg,pam,pbm,pgm,png,pnm,ppm,qoi,tga,tif,tiff,webp";

/// The extensions written to `config.ini` on first run: the drawings a hover is expected to meet.
///
/// `svg` and `svgz` are the documents the browser engine draws, and both spellings of the same
/// format; `wmf` and `emf` are Windows' two metafiles — a list of drawing records that the
/// drawing layer plays back, which is why a preview of one is sharp at any size — and `eps`,
/// `epsi`, `epsf`, `epi`, `ept`, `ept2` and `ept3` are the spellings of an encapsulated
/// PostScript file, read here for the preview picture a writer leaves inside it rather than for
/// the PostScript itself: nothing in this app interprets PostScript, and an ImageMagick that
/// could draw one would still need a Ghostscript installed beside it.
pub const DEFAULT_VECTOR_EXTENSIONS: &str = "emf,epi,eps,epsf,epsi,ept,ept2,ept3,svg,svgz,wmf";

/// The built-in vector list as it stood before `svg` and `svgz` were added to it — which is to
/// say the list the kind had when it was written.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the two documents would be left out of every `config.ini` already written, and a
/// list that differs is otherwise the user's own (see `repair_older_lists`).
pub const VECTOR_EXTENSIONS_BEFORE_SVG: &str = "emf,eps,epsi,wmf";

/// The built-in vector list as it stood before the other spellings of an encapsulated
/// PostScript file were added to it: `epsf`, `epi` and `ept` with `ept2` and `ept3`.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the five spellings would reach a fresh installation only: every `config.ini`
/// already written holds the list as it was, and a list that differs is otherwise the user's own
/// (see `repair_older_lists`).
pub const VECTOR_EXTENSIONS_BEFORE_THE_EPS_SPELLINGS: &str = "emf,eps,epsi,svg,svgz,wmf";

/// The extensions written to `config.ini` on first run: the layered documents and the project
/// containers a hover is expected to meet.
///
/// `psd` and `psb` are Photoshop's two, read for the merged picture the format keeps at the end
/// of the file rather than for the layers; `kra` and `ora` are Krita's and OpenRaster's, which
/// are zip containers holding that same picture as a file of its own; `procreate` is
/// Procreate's, a container of the same kind; and `sketch`, `fig` and `xd` are containers of that
/// kind as well. What each one is read for is in `psd_image` and `project_image`, and a container
/// none of them can open is a file that shows no preview, like any other format this app has no
/// reader for.
///
/// `cdr` is deliberately *not* here, though a CorelDRAW document is a design document. What this
/// app could read of one by itself is the picture CorelDRAW keeps for a file browser — 96 to 256
/// pixels across — and a preview of that is worse than no preview at all: what the drawing is, is
/// in the `[libre]` list, drawn by LibreOffice where it is installed, and a machine without it
/// shows a `.cdr` no preview rather than a blurred thumbnail. The name sits in exactly one list so
/// that this cannot come back by accident; see `libre_formats`.
///
/// `ai` is the one name here that is usually not this kind's at all: an Illustrator document saved
/// with `Create PDF Compatible File` is a PDF, and the PDF gate claims it before this list is
/// asked. What reaches here under that name is a document saved without that compatibility, which
/// is an encapsulated PostScript file — the artwork as a program, with the preview an older
/// Illustrator left beside it — and what a preview of one can be is what that preview holds. A
/// document that carries none shows nothing, which is the answer every file with no reader gets.
pub const DEFAULT_DESIGN_EXTENSIONS: &str = "ai,fig,kra,ora,procreate,psb,psd,sketch,xd";

/// The built-in design list as it stood while `cdr` was an entry of it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written,
/// which is what takes the name out of every `config.ini` already written. The engine draws a
/// CorelDRAW document now and this list does not; see `libre_formats`.
pub const DESIGN_EXTENSIONS_WITH_CDR: &str = "ai,cdr,fig,kra,ora,procreate,psb,psd,sketch,xd";

/// The built-in design list as it stood before the two drawing applications that write a
/// container of their own — `cdr`, CorelDRAW's, and `procreate`, Procreate's — were added to it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the two names would reach a fresh installation only: every `config.ini` already
/// written holds a list that differs from the built-in one, and a list that differs is otherwise
/// the user's own (see `repair_older_lists`).
pub const DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE: &str = "ai,fig,kra,ora,psb,psd,sketch,xd";

/// The built-in design list as it stood before `ai` was added to it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written.
/// Without that, the entry would reach a fresh installation only: every `config.ini` already
/// written holds a list that differs from the built-in one, and a list that differs is otherwise
/// the user's own (see `repair_older_lists`).
pub const DESIGN_EXTENSIONS_BEFORE_AI: &str = "fig,kra,ora,psb,psd,sketch,xd";

/// The extensions written to `config.ini` on first run: the font formats a hover is expected to
/// meet.
///
/// The two webfont containers and the three desktop ones, which are the same outlines in
/// different wrappers. What draws one is the browser engine the SVG previews already use — it
/// reads all five, and a `.ttc` is the one of them it cannot be *pointed* at, which is why that
/// face is written out as a font of its own on this side; see `webview_preview` and
/// `font_preview`. What this side reads of them is two tables, the character map and the name, and
/// the drawing is the engine's.
///
/// Those five are also the *whole* of what the engine can be given, which is why no other font
/// format is listed here and none can be added that would work. Every font a page asks for goes
/// through the engine's own sanitizer before its font stack sees it — the OpenType Sanitizer,
/// which parses OpenType in its two shapes and the two WebFont containers and turns everything
/// else down — so a PostScript Type 1 face (`pfa`, `pfb`), a Macintosh suitcase (`dfont`), a
/// Windows bitmap font (`fon`) and a face inside a collection's other shape are all files the
/// specimen can be told about and cannot be drawn with. Measured on the runtime this app uses: a
/// `.ttf` loads and an 11 KB `.fon` beside it is refused, which is the sanitizer answering rather
/// than the font.
pub const DEFAULT_FONT_EXTENSIONS: &str = "otf,ttc,ttf,woff,woff2";

/// The extensions written to `config.ini` on first run: every sound this app has an engine to
/// play, in the one list, because which engine plays a format is the machine's answer and not
/// the list's.
///
/// The native families are here — WAV, MP3, AAC in its two spellings, WMA, FLAC, ALAC, the AMR
/// pair, AIFF, DSD's two containers, AC-3 and DTS — and so is everything an installed FFmpeg
/// reads and Windows does not: Ogg Vorbis and Opus, Matroska's audio, Musepack, WavPack, Monkey's
/// Audio, True Audio, Shorten, TAK, OptimFROG, CAF, AU, VOC and RealAudio. A name whose decoders
/// are on neither engine is a name that shows nothing, and the list is a list of the ones that do.
pub const DEFAULT_AUDIO_EXTENSIONS: &str = "aac,ac3,aif,aifc,aiff,amr,ape,au,awb,caf,dff,dsf,dts,dtshd,eac3,flac,m4a,m4b,mka,mp2,mp3,mpa,mpc,oga,ogg,ofr,ofs,opus,ra,shn,snd,spx,tak,tta,voc,wav,wave,wma,wv";

/// The extensions written to `config.ini` on first run under `[video]`: the containers and raw
/// streams the codecs Windows 11 itself has can demux and decode — the ISO base media family, AVI,
/// ASF, Matroska and WebM, and the MPEG-1, MPEG-2 and MPEG-4 elementary, program and transport
/// streams. It is the same list the README's *Videos Windows 11 plays itself* row names, and the
/// two are meant to be read side by side: what is here is what a machine with no FFmpeg on it
/// still plays.
///
/// Which list a name is in decides which engine a video is played by, but only on a machine with
/// no FFmpeg on it: where FFmpeg is installed its player plays every video there is, both lists or
/// neither, and where it is not installed this list is the whole of the question. A file of these
/// names is then played by the media engine Windows has, in this app's own window, so its frames
/// are this app's to draw — a pinned one is resized by its edges, maximized by its caption and
/// dragged by its picture, and its transport bar is a real control. A name in `[ffmpeg]` beside it
/// has no player here at all, because there is nothing to fall back to and the engine is not asked
/// about it; that is what the list has always meant, and it is a file with no preview rather than
/// a broken one.
///
/// What is *not* claimed here is that the machine in hand decodes every file of one of these
/// names: a `.mkv` of HEVC on a machine with no HEVC codec, or a `.mp4` of ProRes, is a file the
/// engine is asked about and turns down. That question is asked of the engine itself, once per
/// file and version (`video_player::plays`), and it is asked only where the engine can still be
/// the answer — a file it turns down is one nothing plays here, and where FFmpeg is installed it
/// was never going to be asked about at all (see `preview_window::route_video`).
pub const DEFAULT_VIDEO_EXTENSIONS: &str = "3g2,3gp,3gpp,asf,avi,dvr-ms,m1v,m2t,m2ts,m2v,m4v,mkv,\
mov,mp4,mpe,mpeg,mpg,mts,qt,ts,vob,webm,wmv";

/// The extensions written to `config.ini` on first run under `[ffmpeg]`: every container and raw
/// video stream FFmpeg is able to demux that Windows' own codecs are not asked about — the streams
/// no decoder of Windows' reads (AV1, VC-1, Flash's formats), the containers whose handler
/// Windows does not ship (RealMedia, MXF, NUT, the game and camera formats), and the names the ISO
/// base media family is shared with where what is inside is not what Windows decodes.
///
/// It is the README's *Needs FFmpeg* list, and it is the rest of the one list the two were. A name
/// here is played by FFmpeg's `ffplay` wherever FFmpeg is installed, which is every machine that
/// has it, and the two lists do not differ there at all — so this is what the list is for on the
/// machines it is written for: a machine with no FFmpeg installed shows nothing for one of these
/// names, rather than asking the engine about a format nothing here will play. Moving a name from
/// here to `[video]` is the whole of asking the media engine about it instead, and there is one
/// cost worth stating rather than discovering: a name the engine cannot open has no preview at all
/// on such a machine, because the probe that decides it is the last thing standing between the
/// file and a player (see `DEFAULT_VIDEO_EXTENSIONS` and `preview_window::route_video`).
pub const DEFAULT_FFMPEG_EXTENSIONS: &str =
    "264,265,266,apv,av1,avc,avs,avs2,avs3,bik,bk2,c93,cavs,cdg,cdxl,cin,cpk,dav,\
dif,divx,drc,dv,evc,f4v,flm,flv,gxf,h261,h263,h264,h265,h266,h26l,hevc,ifv,imx,ismv,ivf,ivr,\
kux,m2p,mj2,mjpeg,mjpg,mk3d,moflex,mpv,mve,mvi,mxf,mxg,nsv,nut,obu,ogm,ogv,pmp,psp,rcv,rm,rmvb,\
roq,rsd,smk,str,swf,thp,tod,tp,tr,ty,ty+,usm,vc1,vc2,viv,vro,vvc,vw,wtv,xl,xmv,y4m,yop";

/// The one list the two above were one list of: every container and raw video stream FFmpeg is
/// able to demux, which is what every build before the split wrote under `[video]`.
///
/// It is here for the same reason the lists beside the other kinds are: a list is only ever read
/// out of `config.ini` — nothing in the tray edits one — so a `[video]` list holding exactly these
/// entries is this app's own older list rather than an edit somebody made by hand, and it is what
/// tells the repair that a file written before the split is to be split rather than kept whole. A
/// list with any entry added, removed or spelled differently is the user's and is left exactly as
/// it is.
pub const VIDEO_EXTENSIONS_BEFORE_THE_SPLIT: &str = "264,265,266,3g2,3gp,3gpp,apv,asf,av1,avc,avi,avs,avs2,avs3,bik,bk2,c93,cavs,cdg,cdxl,cin,cpk,dav,\
dif,divx,drc,dv,dvr-ms,evc,f4v,flm,flv,gxf,h261,h263,h264,h265,h266,h26l,hevc,ifv,imx,ismv,ivf,\
ivr,kux,m1v,m2p,m2t,m2ts,m2v,m4v,mj2,mjpeg,mjpg,mk3d,mkv,moflex,mov,mp4,mpe,mpeg,mpg,mpv,mts,mve,\
mvi,mxf,mxg,nsv,nut,obu,ogm,ogv,pmp,psp,qt,rcv,rm,rmvb,roq,rsd,smk,str,swf,thp,tod,tp,tr,ts,ty,\
ty+,usm,vc1,vc2,viv,vob,vro,vvc,vw,webm,wmv,wtv,xl,xmv,y4m,yop";

/// The extensions written to `config.ini` on first run: the archive formats a hover is expected to
/// meet, plus the zip containers that cost nothing extra to list (`.jar`, `.apk`, `.xpi`, `.cbz`
/// are all zips).
///
/// `tar.gz` is a name rather than an extension — the last dot of `sources.tar.gz` is `gz`, which is
/// not an archive on its own — so an entry containing a dot is matched against the end of the
/// file's name instead.
pub const DEFAULT_ARCHIVE_EXTENSIONS: &str = "7z,apk,jar,rar,tar,tar.gz,tgz,xpi,zip,zipx";

/// The built-in `[archive]` list as it stood while `cbz` was an archive rather than a comic.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has typed it — so it is brought up to the built-in list rather than kept as written,
/// which is what takes the name out of every `config.ini` already written. What it costs is the
/// page of contents a `.cbz` used to be shown as: the name is a comic's now, and what a hover on
/// one shows is its first page (see `ebook_formats` and `comic_preview`). A user who would rather
/// have the listing back adds `cbz` to this list again — both lists are theirs — and the order the
/// two are asked in is what decides; see `content_type::kind_claiming`.
///
/// An archive this app still reads itself, in the order it was written then: the dotted `tar.gz`
/// is what a tarball is claimed by, so it is written as it was.
pub const ARCHIVE_EXTENSIONS_BEFORE_THE_COMICS: &str =
    "7z,apk,cbz,jar,rar,tar,tar.gz,tgz,xpi,zip,zipx";

/// The extensions written to `config.ini` on first run: the Word, Excel and PowerPoint formats a
/// hover is expected to meet, templates and slide shows included.
pub const DEFAULT_OFFICE_EXTENSIONS: &str =
    "doc,docm,docx,dot,dotm,dotx,pot,potm,potx,pps,ppsm,ppsx,ppt,pptm,pptx,xls,xlsb,xlsm,xlsx,xlt,\
xltm,xltx";

/// The extensions written to `config.ini` on first run: the PDF's own three spellings, and the
/// three comic containers this app reads a first page out of.
///
/// The PDF's spellings are the first three letters of the family and the ones that were never a
/// list's before: `pdf` is the format, `pdfa` the archival profile of it, and `epdf` the
/// encapsulated one — all three drawn by the same reader, so all three are here, or a file named
/// with the two the format's world writes beside the first would stop being previewed at all.
///
/// The comics are `cbz`, a zip of plates, `cbr`, a rar of them, and `cbc`, which is Calibre's own
/// container: a zip whose entries are the pages of several comics under a folder each, with a
/// `comics.txt` naming them. Nothing invented any of them as a drawing format — each is a box —
/// and all three are read here rather than handed to an engine, because the engine that reads
/// comics unpacks them, decodes every plate, rewrites it and builds a document out of the whole
/// thing: measured against the comics this was built for, a hundred megabytes of plates is minutes
/// of work for a preview that is one page, against milliseconds for reading that one page out of
/// the box (see `comic_preview`).
///
/// Deliberately absent: the books the ebook engine converts (`azw`, `azw3`, `azw4`, `djvu`, `epub`,
/// `fb2`, `htmlz`, `lrf`, `mobi`, `pml`, `prc`, `snb`, `tcr`, `chm`, `lit`), which are the
/// `[calibre]` list's because none of them is a container this app can read itself, and every name
/// the other lists carry, because a name sits in exactly one list.
pub const DEFAULT_EBOOK_EXTENSIONS: &str = "cbc,cbr,cbz,epdf,pdf,pdfa";

/// The extensions written to `config.ini` on first run: the document formats the engine reads
/// that this app has no reader of its own for.
///
/// Nothing here is a picture, a drawing, a video, an archive, a font or a text file — those are
/// this app's own kinds — and nothing here is a PDF. What is here is what those kinds leave: the
/// word processors that came before the modern one and the ones beside it (`wpd`, `wps`, `abw`,
/// `lwp`, `cwk`, `hwp`, `602`, `wri`), the older spreadsheets (`123`, `wk1`, `wk3`, `wk4`, `wks`,
/// `slk`, `dif`, `dbf`, `wb2`, `wq1`, `wq2`, `gnumeric`, `xlw`), the presentations (`sda`, `sdc`,
/// `sdd`, `sdw`, `sxi`, `sti`), the drawings whose own format this app does not read (`cdr`, `cmx`,
/// `dxf`, `wpg`, `pub`, `zmf`, `cgm`, `pct`, `met`, `svm`, and the Visio names the engine's
/// filter declares) — and the open formats themselves, `odt`, `ods`, `odp`, `odg`, `odc`, `odb`,
/// `odf` and their friends, which no other list claims and which the engine reads exactly.
///
/// What a name is asked about is settled by the engine's own filter registry rather than by the
/// list of formats the engine is *said* to support, and those two are not the same list. A name
/// belongs here only where a filter that imports declares that very extension — read one name at a
/// time out of an installed engine's `share/registry/*.xcd` — because a name no filter declares
/// is a launch that answers nothing: what the engine falls back to is the file's own content,
/// where a filter it cannot use may spin rather than answer (see `swf`), and where it answers at
/// all it answers with nothing. Four groups an earlier list held came out that way:
///
/// * `qxp` is the older QuarkXPress document. The filter reads `qxd` and `qxt`.
/// * `pm3`, `pm4` and `pm5` are PageMaker before 6. The filter reads `pm`, `p65`, `pm6` and `pmd`.
/// * `vssm`, `vst`, `vstm`, `vtx` and `vsx` are Visio stencils and templates. The filter reads
///   `vdx`, `vsd`, `vsdm`, `vsdx` and `vstx`.
/// * `epub` is a name the engine *writes* rather than reads — the one filter that declares it is
///   an export filter, which is the wrong direction for a preview — and `agd`, `fhd`, `jtd`,
///   `jtt`, `plt`, `pxl`, `rl`, `sdp`, `sgf`, `sgl`, `uof`, `uop`, `uos`, `uot` and `vor` are
///   declared by no filter at all.
///
/// A machine whose engine does read one of those can have the name back by adding it: the list is
/// what the user edits, and a name added to it is asked about from the next read.
///
/// The engine imports a few names that are deliberately not here as well, because there is
/// nothing in them to preview — each is a file *about* a document rather than one:
///
/// * `ase` and `gpl` are colour palettes. LibreOffice reads them to fill a colour picker, and
///   what a page would be drawn from one is nothing at all.
/// * `oxt` is an extension package: a zip of the files that install something into the engine,
///   which is not a document any more than a `.zip` is.
/// * `smf` means StarMath to LibreOffice and a MIDI sequence to everything else that reads the
///   name, and a hover cannot tell which one it has. A MIDI file claimed as a document would be a
///   launch that answers nothing, every time, for a name that is not this app's.
/// * `kth` is a Keynote theme and `iqy` is a web query: the first is a preset, the second is a
///   line of text naming a URL, and neither is a document.
/// * `swf` is a Flash animation rather than a document, and the engine does not draw one: asked
///   to convert one, its filter chain spins with a core at a hundred percent and never writes a
///   page — measured on real files, and past every bound a conversion is given. It was in this
///   list once, and a file of that name cost a launch and a core for as long as the engine was
///   left to it. What reads a Flash file is FFmpeg's own SWF demuxer — the drawings, the sounds
///   and the timeline of one — so the name is in the video list, which is where it belongs, and is
///   not here (see `video_formats`).
pub const DEFAULT_LIBRE_EXTENSIONS: &str = "123,602,abw,cdr,cgm,cmx,cwk,dbf,dif,dxf,fodg,fodp,fodt,gnm,gnumeric,hwp,key,lwp,mcw,met,mw,numbers,odb,odc,odf,odg,odm,odp,ods,odt,oth,otg,otm,otp,ots,ott,pages,pcd,pct,pcx,pdb,pm6,pmd,psw,pub,ras,sda,sdc,sdd,sdw,slk,stc,std,sti,stw,svm,sxd,sxg,sxi,sxm,sxw,vdx,vsd,vsdm,vsdx,vstx,wb2,wk1,wk3,wk4,wks,wpg,wq1,wq2,wpd,wps,wri,xlw,zabw,zmf";

/// The built-in `[libre]` list as it stood before the names the engine cannot read were taken out
/// of it: the one `swf` was in, and the four groups above with it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit —
/// nobody has touched it — so it is brought up to the built-in list rather than kept as written,
/// which is what takes those names out of every `config.ini` already written. What each of them
/// cost is written beside it in the list above: a launch that answered nothing, and for `swf` a
/// launch that never ended at all.
pub const LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ: &str = "123,602,abw,agd,cdr,cgm,cmx,cwk,dbf,dif,dxf,epub,fhd,fodg,fodp,fodt,gnm,gnumeric,hwp,jtd,jtt,key,lwp,mcw,met,mw,numbers,odb,odc,odf,odg,odm,odp,ods,odt,oth,otg,otm,otp,ots,ott,pages,pcd,pct,pcx,pdb,plt,pm3,pm4,pm5,pm6,pmd,psw,pub,pxl,qxp,ras,rl,sda,sdc,sdd,sdp,sdw,sgf,sgl,slk,stc,std,sti,stw,svm,swf,sxd,sxg,sxi,sxm,sxw,uof,uop,uos,uot,vdx,vor,vsd,vsdm,vsdx,vssm,vst,vstm,vstx,vtx,vsx,wb2,wk1,wk3,wk4,wks,wpg,wq1,wq2,wpd,wps,wri,xlw,zabw,zmf";

/// The extensions written to `config.ini` on first run: the picture formats ImageMagick reads and
/// this app has no reader of its own for.
///
/// The camera raw formats come first in spirit if not in order — `3fr`, `arw`, `cr2`, `cr3`,
/// `crw`, `dcr`, `dng`, `erf`, `fff`, `iiq`, `k25`, `kdc`, `mdc`, `mef`, `mos`, `mrw`, `nef`,
/// `nrw`, `orf`, `pef`, `raf`, `raw`, `rmf`, `rw2`, `rwl`, `sr2`, `srf`, `srw` and `x3f`, from
/// the raw decoder the engine is built with, which is every raw format a camera writes that is
/// still met with — and they are what this list is mostly about: nothing else on a Windows
/// machine opens one, the shell shows a thumbnail from the picture the camera left inside the file
/// and nothing else, and a hover onto one shows nothing at all rather than the photograph.
///
/// What follows them, in the order the module documentation is written in:
///
/// * The formats the engine has a coder of its own for and no other list claims: the medical
///   scanner's `dcm`, the film scanner's `dcx` and `dpx`, the astronomer's `fit`, `fits` and
///   `fts`, the compositor's `j2c`, `j2k`, `jp2`, `jpc`, `jpm` and `jpt`, the animator's `jng` and
///   `mng`, the illustrator's `xcf`, `xbm` and `xpm`, the painter's `sgi`, the engineer's `vicar`,
///   the phone's `wbmp`, the cursor's `cur`, and the engine's own `miff`.
/// * The formats the registry declares read support for that no preview was ever asked for: a
///   pixel-art program's sprites (`ase`, `aseprite`), a floppy-era paint program's pictures
///   (`mac`, `pix`, `rla`, `rle`, `art`, `cut`, `wbinfo`), a fax machine's pages (`fax`, `g3`,
///   `g4`), a scanner's `pgx`, a texture of a console's (`tim`, `tm2`), an icon of a robot's
///   (`rgf`), an embroidery machine's pattern (`pes`), a spectrum analyser's screen (`scr`), a
///   telescope's or a satellite's frame (`hrz`, `ipl`, `fl32`, `sct`, `jnx`), a colour lookup
///   table (`cube`), a document's provenance record (`c2pa`), a markup language (`pango`), and the
///   rest of the names beside them. Each is a picture the engine really draws; what none of them
///   has is a reason to have been picked over the others, and the list shipping with them is what
///   settles that.
/// * The second spelling of a format whose first spelling another list holds — `pict`, `sun`,
///   `pcds`, `dxt1` and `dxt5`, `icb`, `vda` and `vst`, `picon` — which the engine reads and the
///   engine the other spelling belongs to does not (see the module documentation).
/// * The raw sample dumps (`rgb`, `rgba`, `gray`, `cmyk` and their kin, and the `group4` fax
///   bitstream), whose shape comes from their own length: a dump whose length does not settle one
///   is a file this app shows nothing for, and one that does is read at the size that came out of
///   it.
///
/// What is deliberately *not* here is named in the module documentation: the alternative spelling
/// of a name another list already claims, the fonts, the documents that need a delegate, the
/// engine's own notation, and `xwd`, which this build ships no coder for.
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
/// list as it was (see `repair_older_lists`).
pub const MAGICK_EXTENSIONS_BEFORE_THE_REST: &str = "3fr,arw,cr2,cr3,crw,cur,dcm,dcr,dcx,dng,dpx,erf,fff,fit,fits,fts,iiq,j2c,j2k,jng,jp2,jpc,jpm,jpt,k25,kdc,mdc,mef,miff,mng,mos,mrw,nef,nrw,orf,pef,pfm,raf,raw,rmf,rw2,rwl,sgi,sr2,srf,srw,vicar,wbmp,x3f,xbm,xcf,xpm";

/// The extensions written to `config.ini` on first run: the names the tools of an installed
/// PeaZip can be asked about, that no list of this app's own already claims, and that are
/// containers of files rather than programs.
///
/// Three groups, in the order they are written:
///
/// * The archives and installers nothing else on the machine opens: `001`, `ar`, `arc`, `arj`,
///   `cab`, `chm`, `cpio`, `deb`, `esd`, `hfs`, `hfsx`, `hxs`, `iso`, `lha`, `lzh`, `msi`, `msp`,
///   `pkg`, `ppkg`, `rpm`, `swm`, `udf`, `wim`, `xar`, `xip` and `zpaq`. Some are containers of
///   files in the ordinary sense (an installer, a compiled help file, a Linux package, a disk
///   image), some are the volume of a backup, and all of them are read by a tool of PeaZip's and by
///   nothing this app has. Two of them are that tool's rather than the console archiver's: an
///   `arc` is FreeArc's and a `zpaq` is zpaq's, and neither is read by the archiver at all. A
///   Microsoft Reader book was of this group once and is not any more: the ebook engine draws a
///   page of one, so it is `[calibre]`'s, and `PEAZIP_EXTENSIONS_BEFORE_THE_EBOOKS` below is what
///   takes it out of a file this app wrote before that. A compiled help file went the other way —
///   out of this group and back into it — which is what `PEAZIP_EXTENSIONS_BEFORE_THE_HELP_FILE` is
///   written down for.
/// * The single-stream compressors: `bcm`, `br`, `bz2`, `bzip2`, `gz`, `gzip`, `lpaq8`, `lzma`,
///   `xz`, `z` and `zst`, with the tarball spellings that name them (`taz`, `tbz`, `tbz2`, `tpz`,
///   `tzst`). One of these is a file put through a compressor rather than a container, so the
///   answer is one member, often one whose name is not in the stream at all — and the reading of
///   that answer is what puts a name to it (see `archive_listing`). They are here because a tool of
///   PeaZip's reads them and this app does not, and because what a hover shows for one — the name,
///   the size where the tool knows it, and how much smaller it was made — is the useful part of
///   opening it. `bcm`, `br` and `lpaq8` are the three whose tools cannot say even that much;
///   nothing is started for one of them.
/// * And the disk images whose names are their own: `apfs`, `cramfs`, `dmg`, `qcow`, `qcow2`,
///   `squashfs`, `vdi`, `vhd`, `vhdx`, `vmdk`.
///
/// Deliberately absent, and each for a reason the module documentation gives: the names another
/// list of this app's already reads (`7z`, `zip`, `rar`, `tar`, `zipx`, `jar`, `apk`, `xpi`, `cbz`,
/// `cbr`, `cbc`, `chm`, `lit`, `tgz`, `pmd`, `swf`, `flv`, `doc`, `xls`, `ppt`), the programs the
/// engine lists as resources (`exe`, `dll`, `sys`, `obj`, `elf`, `macho`, `te`, `b64`, `ihex`,
/// `simg`, `uefif`, `scap`, `lpimg`, `nsis`, `mslz`, `mub`), the extensions that are words rather
/// than formats (`img`, `ext`, `ext2`, `ext3`, `ext4`, `fat`, `ntfs`, `apm`, `mbr`, `gpt`), the
/// codecs its build carries no format for and no tool of its own opens (`lz4`, `lz5`, `lizard`,
/// `flzma2`), the two names this installation carries no tool for at all (`pea`, which no tool of
/// PeaZip's lists, and `paq8`, whose folder holds no executable — both are written down in
/// TODO.md), and the compression formats this app has no need of an engine for (`xz` is here,
/// `lzma86` and `base64` are not).
pub const DEFAULT_PEAZIP_EXTENSIONS: &str =
    "001,apfs,ar,arc,arj,bcm,br,bz2,bzip2,cab,chm,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lpaq8,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zpaq,zst";

/// The built-in `[peazip]` list as it stood while a compiled help file was the ebook engine's.
///
/// A file holding exactly these entries is the app's own earlier list rather than a user's edit,
/// so it is brought up to the built-in list rather than kept as written — which is what gives the
/// name back to the archiver, on an installation that ran the build that had taken it away. `chm`
/// is the only name to have moved twice, and the reason it moved back is worth putting beside it:
/// what the engine draws for one is a page and takes two to three seconds to draw, and what a help
/// file is hovered for is usually nothing at all — so the listing that is there immediately is the
/// better answer, and the page it gives up is one nobody was waiting for (see `calibre_formats`).
pub const PEAZIP_EXTENSIONS_BEFORE_THE_HELP_FILE: &str =
    "001,apfs,ar,arc,arj,bcm,br,bz2,bzip2,cab,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lpaq8,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zpaq,zst";

/// The built-in `[peazip]` list as it stood while `chm` and `lit` were the archiver's names.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit, so
/// it is brought up to the built-in list rather than kept as written — which is what takes the two
/// names out of every `config.ini` already written. Both are books rather than archives: a compiled
/// help file and a Microsoft Reader book are read by the ebook engine and previewed as a page of
/// one, which needs Calibre installed, where the archiver listed what they hold (see
/// `calibre_formats`). A user who would rather have the listing back adds the name to this list
/// again, and takes it out of `[calibre]` with it.
pub const PEAZIP_EXTENSIONS_BEFORE_THE_EBOOKS: &str =
    "001,apfs,ar,arc,arj,bcm,br,bz2,bzip2,cab,chm,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lit,lpaq8,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zpaq,zst";

/// The list this app wrote before the tools beside the console archiver were driven: the names that
/// list held, which is what tells a file written by that build from one a user has edited (see
/// `repair_older_lists`).
///
/// A list holding exactly these entries is this app's own — nobody typed it — and is brought up to
/// `DEFAULT_PEAZIP_EXTENSIONS`, which is how an installation that already exists is given the names
/// the archiver's own table never declared: `arc`, `zpaq`, `br`, `bcm` and `lpaq8`. The entries are
/// written in the order they were written then, since that is what a file holds.
pub const PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS: &str =
    "001,apfs,ar,arj,bz2,bzip2,cab,chm,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lit,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zst";

/// The extensions written to `config.ini` on first run: the ebook formats the engine reads as input
/// that no list of this app's own already claims.
///
/// The four groups are the module documentation's, in the order they are written there: the Kindle
/// and Mobipocket family (`azw`, `azw3`, `azw4`, `mobi`, `prc`), the two books whose text is packed
/// the way no reader here unpacks it (`chm`, `lit`), the open and single-reader formats (`djvu`,
/// `epub`, `fb2`, `lrf`), and the formats of the dedicated readers and of the engine itself
/// (`htmlz`, `pml`, `snb`, `tcr`).
///
/// Deliberately absent, and each for a reason the module documentation above gives: the names
/// another list of this app's already reads (`cbz`, `docx`, `odt`, `pdb`, `html`, `rtf`, `txt`,
/// `pdf`), the comic books, which this app reads itself and needs no engine for (`cbr`, `cbc`), the
/// name that is a programming language (`rb`), the Sony container's protected spelling (`lrx`), and
/// the names the engine does not read at all (`tpz`, and the `kfx` a plugin would be needed for).
pub const DEFAULT_CALIBRE_EXTENSIONS: &str =
    "azw,azw3,azw4,djvu,epub,fb2,htmlz,lit,lrf,mobi,pml,prc,snb,tcr";

/// The built-in `[calibre]` list as it stood while a compiled help file was this engine's to draw.
///
/// `chm` was in this list for one build and is not any more: the page the engine draws for one is
/// right, and the two to three seconds it takes to draw it is not — a help file is a file a pointer
/// crosses on its way somewhere else, and the listing the archiver prints for one is there before a
/// hover has finished settling (see `peazip_formats`, which is where the name is again). A file
/// holding this list is a file this app wrote, so it is brought up to the list of now rather than
/// kept as written, and the name goes back where it came from.
pub const CALIBRE_EXTENSIONS_WITH_THE_HELP_FILE: &str =
    "azw,azw3,azw4,chm,djvu,epub,fb2,htmlz,lit,lrf,mobi,pml,prc,snb,tcr";

/// The built-in `[calibre]` list as it stood before `chm` and `lit` became the engine's names.
///
/// A file holding exactly these entries is this app's own earlier list rather than a user's edit —
/// nobody has typed it — so it is brought up to the built-in list rather than kept as written,
/// which is what gives an installation that already exists the name that stayed. Without it that
/// name would reach a fresh installation only: every `config.ini` already written holds the list as
/// it was, and a list nobody has touched is indistinguishable from one a user edited unless the
/// older spellings of it are written down here (see `repair_older_lists`).
///
/// Both names were the `[peazip]` list's until then, so the same change takes them out of that
/// list. `chm` has since gone back — see [`CALIBRE_EXTENSIONS_WITH_THE_HELP_FILE`] — so a file can
/// hold either of the two older spellings of this list, and both are written down for that reason: a
/// list this app shipped is a list it brings up to now, whichever of them it is.
pub const CALIBRE_EXTENSIONS_BEFORE_THE_EBOOKS: &str =
    "azw,azw3,azw4,djvu,epub,fb2,htmlz,lrf,mobi,pml,prc,snb,tcr";

/// Extensions previewed as text before the list is edited in `config.ini`.
///
/// The list is written to the configuration on first run and read back from there, so adding or
/// removing an extension is an edit in the file rather than a rebuild. Anything in it that no syntax
/// definition claims is still previewed, as plain text.
///
/// `.ts` and `.mts` belong to both lists: they are TypeScript sources and MPEG transport streams,
/// so the video gate decides those two by content (an MPEG-TS sync byte) and a TypeScript file falls
/// through to a text preview. Every other extension here is disjoint from the image, video and PDF
/// gates.
pub const DEFAULT_TEXT_EXTENSIONS: &str =
    "adb,adoc,ads,asciidoc,asm,asp,aspx,astro,awk,bash,bat,bib,bzl,c,cc,cfg,cg,cjs,clj,cljc,cljs,\
cmake,cmd,comp,conf,cpp,cs,csh,cshtml,css,csv,csx,cts,cxx,d,dart,diff,diz,edn,ejs,el,elm,env,erb,\
erl,ex,exs,f,f03,f77,f90,f95,fish,for,frag,fs,fsi,fsx,ftn,fx,geom,glsl,go,gql,gradle,graphql,\
groovy,h,haml,hbs,hcl,hh,hlsl,hpp,hrl,hs,htm,html,hxx,inc,ini,ipynb,java,jl,js,json,json5,jsonc,\
jsonl,jsp,jsx,ksh,kt,kts,latex,less,lhs,liquid,lisp,ll,lock,log,lsp,lua,m,mak,man,markdown,md,\
mdown,metal,mjs,mk,mkd,ml,mli,mm,mts,mustache,nasm,nfo,nim,ninja,nix,njk,org,pas,patch,php,phtml,\
pl,plist,pm,properties,proto,ps1,psd1,psm1,py,pyi,pyw,r,rake,rb,rkt,rmd,rs,rst,rtf,s,sass,scala,\
scm,scss,sh,slim,sol,sql,srt,ss,styl,sv,svelte,svh,swift,tcl,tex,text,tf,tfvars,toml,ts,tsv,tsx,\
twig,txt,v,vbs,vert,vhd,vhdl,vtt,vue,wat,wgsl,xhtml,xml,xsd,xsl,xslt,yaml,yml,zig,zsh";

/// File names previewed as text before the list is edited in `config.ini`.
///
/// An extension is not enough for the files a repository is recognized by: a `.gitignore` has no
/// extension at all as far as the path is concerned — the dot is the start of its *name* — and
/// `LICENSE`, `Makefile` and `Dockerfile` have no dot in them anywhere. So the gate has a second
/// list, of names, and a file matches if either list does.
///
/// The entries are ordered by the name each one matches, a leading dot aside, since a dot begins a
/// name rather than changing it: `.gitattributes` sits where `gitattributes` would, which is also
/// the form this list is written to `config.ini` in and looked up by.
pub const DEFAULT_TEXT_NAMES: &str = "authors,.babelrc,brewfile,caddyfile,changelog,changes,.clang-format,.clang-tidy,cmakelists.txt,\
code_of_conduct,containerfile,contributing,contributors,copying,copyright,dockerfile,\
.dockerignore,.editorconfig,.env,.env.example,.env.local,.eslintignore,.eslintrc,gemfile,\
.gitattributes,.gitconfig,.gitignore,.gitkeep,.gitmodules,gnumakefile,.golangci.yml,history,\
.htaccess,install,jenkinsfile,justfile,licence,license,.mailmap,makefile,makefile.am,makefile.in,\
notice,.npmignore,.prettierignore,.prettierrc,procfile,rakefile,readme,.rustfmt.toml,security,\
.stylelintrc,unlicense,vagrantfile";

/// One configured extension list: where the file writes it, what it holds, what it held before, and
/// which field of the configuration it is.
pub(crate) struct List {
    /// The section of `config.ini` the list is written under.
    ///
    /// One section per list rather than a key of the settings section among fifty others, so the
    /// one long value stays easy to find and edit by hand — and so a section is what a list is keyed
    /// by, which is the only thing about a list two lists of the same app might have in common.
    pub(crate) section: &'static str,
    /// The key within that section: `extensions` for every list but one, which is a list of names
    /// rather than of extensions.
    pub(crate) key: &'static str,
    /// What the list holds on the first run, and what a file whose list is the built-in one is
    /// read back as.
    pub(crate) defaults: &'static str,
    /// The built-in lists this app shipped and then changed.
    ///
    /// Migration data, and the reason it sits in the row rather than beside the list it belongs to:
    /// nothing reads these but the repair below, and a version of a list is a fact about what this
    /// app wrote rather than about what a format is called.
    pub(crate) before: &'static [&'static str],
    /// What an entry of the list may be, which is what the built-in list and every older one is
    /// read by as well — a `tar.gz` was a compound when the list held it first.
    pub(crate) entries: Entries,
    /// Whether the repair that brings a file an older build wrote up to the list of now walks this
    /// row.
    ///
    /// It walks a row this app has shipped an older list under, because a file holding that older
    /// list is a file this app wrote and is brought up to the list of now — and it walks `[ffmpeg]`,
    /// which is new with the split and has no older list, because a file holding the built-in
    /// entries in another order is a list this app wrote whatever order it wrote them in. The six
    /// lists that arrived with their kind are not walked at all: no build of this app ever wrote a
    /// different list under them, so a file that has one has had it edited by a person, and a
    /// person's order is theirs.
    pub(crate) repaired: bool,
    /// What the configuration holds for this list, read: the file is written from it, and the
    /// settings reset is the reset that has to leave it alone.
    pub(crate) held: fn(&AppConfig) -> &[String],
    /// What the configuration is set to for this list, written: a file is read into it, and the
    /// lists reset is the one that puts the built-in list back.
    ///
    /// Two and not one because the two are asked of different borrows — the file is written from a
    /// shared borrow of the whole configuration and no lock of its own, and the read and the two
    /// resets are reached with the only borrow there is.
    pub(crate) set: fn(&mut AppConfig, Vec<String>),
}

impl List {
    /// The built-in list in the form the lookups compare against.
    pub(crate) fn built_in(&self) -> Vec<String> {
        self.entries.sanitize(self.defaults)
    }

    /// The list the configuration holds for this row, which is what a lookup of it is read
    /// against.
    ///
    /// It is the row's own `held` under a name a caller can write, and that is the whole of it:
    /// the alternative is the field of the configuration behind the row, and a list read by field
    /// is a list read from a second place — which is what the sixteen modules this table replaced
    /// were, each holding its list's one read of the configuration and its own copy of the
    /// question.
    pub(crate) fn entries<'a>(&self, config: &'a AppConfig) -> &'a [String] {
        (self.held)(config)
    }

    /// Whether the configuration's list for this row claims `path`.
    ///
    /// What an entry of a list may be and how one of them is compared against a file's name are
    /// the same question asked twice, and for three of the four rules the second follows from the
    /// first: a bare extension, a `C#`-shaped one and a whole file name are each looked up the way
    /// that rule reads them. The fourth is the archive row's, where a compound entry names a whole
    /// file rather than an extension — `tar.gz` — and so is matched against the end of the name as
    /// well as against the extension; that one is asked of `archive_formats`, which is where that
    /// half of the rule is written down.
    ///
    /// It is the row's question and not the router's: nothing here opens the file, nothing is
    /// probed and no half of a kind is settled, so it answers what a name is written in. What kind
    /// a file is, is `routing::kind_of`, which is the same rows asked in an order.
    pub(crate) fn claims(&self, path: &Path, config: &AppConfig) -> bool {
        let entries = self.entries(config);

        match self.entries {
            Entries::Bare | Entries::Hashed => {
                crate::formats::text_formats::matches_configured_extension(path, entries)
            }
            Entries::Compound => crate::formats::archive_formats::claims_in(path, entries),
            Entries::Name => crate::formats::text_formats::matches_configured_name(path, entries),
        }
    }

    /// One list as the file has it: the entries its key names, or the built-in list when the key is
    /// gone — which is a file to write, since what it holds is then not what the app is using.
    ///
    /// What is read here is what the file says: a list anyone has edited keeps its own entries and
    /// its own order, and an empty value is a list with nothing in it rather than a missing one. A
    /// list this app itself wrote and has since added entries to was dealt with before this ran, by
    /// the repair below, so by the time a list is read here it is either the user's or the list of
    /// now.
    pub(crate) fn read(&self, ini: &Ini) -> Vec<String> {
        match ini.get(self.section, self.key) {
            Some(value) => self.entries.sanitize(&value),
            None => self.built_in(),
        }
    }

    /// One list as the file writes it: what the configuration holds, read back by the rule it was
    /// read in by, so a list a caller set by hand is written as the file will read it rather than as
    /// it was given.
    pub(crate) fn write(&self, list: &[String]) -> String {
        self.entries.sanitize(&list.join(",")).join(",")
    }
}

/// The `[image]` list: the still and animated pictures this app decodes itself, so a file refused
/// here is never opened at all.
pub(crate) const IMAGE: List = List {
    section: "image",
    key: "extensions",
    defaults: DEFAULT_IMAGE_EXTENSIONS,
    before: &[
        IMAGE_EXTENSIONS_WITH_SVG,
        IMAGE_EXTENSIONS_BEFORE_DDS,
        IMAGE_EXTENSIONS_BEFORE_SVG,
        IMAGE_EXTENSIONS_BEFORE_CODEC_FORMATS,
        IMAGE_EXTENSIONS_BEFORE_AVCI,
    ],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.image_extensions,
    set: |config, list| config.image_extensions = list,
};

/// The `[vector]` list: the drawings, which a browser or the drawing layer plays rather than a
/// decoder.
pub(crate) const VECTOR: List = List {
    section: "vector",
    key: "extensions",
    defaults: DEFAULT_VECTOR_EXTENSIONS,
    before: &[
        VECTOR_EXTENSIONS_BEFORE_SVG,
        VECTOR_EXTENSIONS_BEFORE_THE_EPS_SPELLINGS,
    ],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.vector_extensions,
    set: |config, list| config.vector_extensions = list,
};

/// The `[design]` list: the layered documents and the project containers holding a flattened
/// picture.
pub(crate) const DESIGN: List = List {
    section: "design",
    key: "extensions",
    defaults: DEFAULT_DESIGN_EXTENSIONS,
    before: &[
        DESIGN_EXTENSIONS_WITH_CDR,
        DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE,
        DESIGN_EXTENSIONS_BEFORE_AI,
    ],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.design_extensions,
    set: |config, list| config.design_extensions = list,
};

/// The `[font]` list: new with its kind, so an older file has no section at all and is given the
/// built-in entries with the key.
pub(crate) const FONT: List = List {
    section: "font",
    key: "extensions",
    defaults: DEFAULT_FONT_EXTENSIONS,
    before: &[],
    entries: Entries::Bare,
    repaired: false,
    held: |config| &config.font_extensions,
    set: |config, list| config.font_extensions = list,
};

/// The `[audio]` list: every sound, in the one list, because which engine plays a format is the
/// machine's answer and not the list's. New with its kind.
pub(crate) const AUDIO: List = List {
    section: "audio",
    key: "extensions",
    defaults: DEFAULT_AUDIO_EXTENSIONS,
    before: &[],
    entries: Entries::Bare,
    repaired: false,
    held: |config| &config.audio_extensions,
    set: |config, list| config.audio_extensions = list,
};

/// The `[video]` list: the containers and streams the media engine Windows has is asked to play,
/// so a pinned window of one is this app's own to draw. It is read only on a machine with no
/// FFmpeg on it; where FFmpeg is installed its player takes every video there is (see
/// `DEFAULT_VIDEO_EXTENSIONS`).
pub(crate) const VIDEO: List = List {
    section: "video",
    key: "extensions",
    defaults: DEFAULT_VIDEO_EXTENSIONS,
    before: &[VIDEO_EXTENSIONS_BEFORE_THE_SPLIT],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.video_extensions,
    set: |config, list| config.video_extensions = list,
};

/// The `[ffmpeg]` list: the rest of the one list `[video]` was, which FFmpeg's player is asked about
/// instead and this app's own window cannot touch. What it means depends on the machine rather
/// than on the list: FFmpeg's player plays both lists wherever it is installed, and this one is
/// the set of names that have nothing to play them at all on a machine where it is not (see
/// `DEFAULT_FFMPEG_EXTENSIONS`).
pub(crate) const FFMPEG: List = List {
    section: "ffmpeg",
    key: "extensions",
    defaults: DEFAULT_FFMPEG_EXTENSIONS,
    before: &[],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.ffmpeg_extensions,
    set: |config, list| config.ffmpeg_extensions = list,
};

/// The `[archive]` list: the one list that holds a name rather than an extension, because a tarball
/// is claimed by the end of a file's whole name.
pub(crate) const ARCHIVE: List = List {
    section: "archive",
    key: "extensions",
    defaults: DEFAULT_ARCHIVE_EXTENSIONS,
    before: &[ARCHIVE_EXTENSIONS_BEFORE_THE_COMICS],
    entries: Entries::Compound,
    repaired: true,
    held: |config| &config.archive_extensions,
    set: |config, list| config.archive_extensions = list,
};

/// The `[office]` list: the documents whose own application may or may not be installed, which is
/// why the engine is asked about a page of one.
pub(crate) const OFFICE: List = List {
    section: "office",
    key: "extensions",
    defaults: DEFAULT_OFFICE_EXTENSIONS,
    before: &[],
    entries: Entries::Bare,
    repaired: false,
    held: |config| &config.office_extensions,
    set: |config, list| config.office_extensions = list,
};

/// The `[ebook]` list: the two readers' worth of names one kind of preview has — the PDF's own three
/// spellings and the comic containers.
pub(crate) const EBOOK: List = List {
    section: "ebook",
    key: "extensions",
    defaults: DEFAULT_EBOOK_EXTENSIONS,
    before: &[],
    entries: Entries::Bare,
    repaired: false,
    held: |config| &config.ebook_extensions,
    set: |config, list| config.ebook_extensions = list,
};

/// The `[libre]` list: the documents the render engine is asked about, which are the ones nothing
/// else here reads.
pub(crate) const LIBRE: List = List {
    section: "libre",
    key: "extensions",
    defaults: DEFAULT_LIBRE_EXTENSIONS,
    before: &[LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.libre_extensions,
    set: |config, list| config.libre_extensions = list,
};

/// The `[magick]` list: the pictures the ImageMagick engine is asked about, which are the camera raw
/// formats above all.
pub(crate) const MAGICK: List = List {
    section: "magick",
    key: "extensions",
    defaults: DEFAULT_MAGICK_EXTENSIONS,
    before: &[MAGICK_EXTENSIONS_BEFORE_THE_REST],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.magick_extensions,
    set: |config, list| config.magick_extensions = list,
};

/// The `[peazip]` list: the archives the console archiver lists and this app has no reader of its
/// own for.
pub(crate) const PEAZIP: List = List {
    section: "peazip",
    key: "extensions",
    defaults: DEFAULT_PEAZIP_EXTENSIONS,
    before: &[
        PEAZIP_EXTENSIONS_BEFORE_THE_HELP_FILE,
        PEAZIP_EXTENSIONS_BEFORE_THE_EBOOKS,
        PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS,
    ],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.peazip_extensions,
    set: |config, list| config.peazip_extensions = list,
};

/// The `[calibre]` list: the books the ebook engine is asked to convert, which is the only way this
/// app can draw a page of one.
pub(crate) const CALIBRE: List = List {
    section: "calibre",
    key: "extensions",
    defaults: DEFAULT_CALIBRE_EXTENSIONS,
    before: &[
        CALIBRE_EXTENSIONS_WITH_THE_HELP_FILE,
        CALIBRE_EXTENSIONS_BEFORE_THE_EBOOKS,
    ],
    entries: Entries::Bare,
    repaired: true,
    held: |config| &config.calibre_extensions,
    set: |config, list| config.calibre_extensions = list,
};

/// The `[text] extensions` key: the extensions read as text, and the one extension list of the
/// sixteen that admits a `#` so that a C# project is a text file.
pub(crate) const TEXT: List = List {
    section: "text",
    key: "extensions",
    defaults: DEFAULT_TEXT_EXTENSIONS,
    before: &[],
    entries: Entries::Hashed,
    repaired: false,
    held: |config| &config.text_extensions,
    set: |config, list| config.text_extensions = list,
};

/// The `[text] names` key: the files a repository is recognized by, which have no extension to match
/// at all.
pub(crate) const NAMES: List = List {
    section: "text",
    key: "names",
    defaults: DEFAULT_TEXT_NAMES,
    before: &[],
    entries: Entries::Name,
    repaired: false,
    held: |config| &config.text_names,
    set: |config, list| config.text_names = list,
};

/// Every list, in the order the file writes them.
///
/// The order is the table's own and nothing reads it back: the file sorts its own sections, and the
/// two things that walk this list in step — the settings reset and the lists reset — are both
/// walking the same one.
pub(crate) const LISTS: &[List] = &[
    ARCHIVE, AUDIO, CALIBRE, DESIGN, EBOOK, FONT, FFMPEG, IMAGE, LIBRE, MAGICK, OFFICE, PEAZIP,
    TEXT, NAMES, VECTOR, VIDEO,
];

/// Every list, read out of the file into the configuration.
///
/// A list is what the file says it is, and a key that is gone is a list the file no longer has: the
/// built-in entries are put back, and the file is written out again because it does not say what the
/// app is using. An empty value is not the same thing — it is a list the user emptied, and it is
/// kept as written.
pub(crate) fn read_all(ini: &Ini, config: &mut AppConfig) {
    for list in LISTS {
        (list.set)(config, list.read(ini));
    }
}

/// Every list, written out of the configuration into the file.
pub(crate) fn write_all(config: &AppConfig, ini: &mut Ini) {
    for list in LISTS {
        ini.set(
            list.section,
            list.key,
            Some(list.write((list.held)(config))),
        );
    }
}

/// Every list the configuration holds, as one piece, so a caller can put them back where they were.
///
/// The two resets are the reason it exists: one puts every setting back at what this build
/// recommends and must leave the lists exactly as they are, the other puts the lists back and must
/// leave everything else alone, and both are one move with this in between. A list added to
/// `AppConfig` later belongs in the table above too — a field no row names is a list nothing writes,
/// and nothing in the tree would say so.
pub(crate) fn held(config: &AppConfig) -> Vec<Vec<String>> {
    LISTS
        .iter()
        .map(|list| (list.held)(config).to_vec())
        .collect()
}

/// Put every list back, in the order [`held`] took them in.
pub(crate) fn put(config: &mut AppConfig, taken: Vec<Vec<String>>) {
    for (list, held) in LISTS.iter().zip(taken) {
        (list.set)(config, held);
    }
}

/// Put the built-in list back, under every name.
pub(crate) fn reset_built_in(config: &mut AppConfig) {
    for list in LISTS {
        (list.set)(config, list.built_in());
    }
}

/// Whether a list holds exactly the entries the built-in list holds, order aside.
fn same_entries(list: &[String], canonical: &[String]) -> bool {
    list.len() == canonical.len() && canonical.iter().all(|entry| list.contains(entry))
}

/// The built-in lists this app shipped and then changed, brought up to the list of now.
///
/// A list is only ever read out of a file — nothing in the tray edits one — so a list that differs
/// from the built-in one is either this app's own older list, written before an entry was added to
/// it or before its entries were put in alphabetical order, or an edit somebody made by hand. The
/// two are told apart by their entries, and only one of them is rewritten: a list holding exactly
/// the entries of a list this app shipped is this app's own — nobody typed it — so it is replaced
/// with the built-in list, while a list with any one entry added, removed or spelled differently is
/// the user's and is kept exactly as it is. Without this, an entry added to a built-in list would
/// reach a fresh installation only, since every file already written holds the list as it was.
///
/// Telling them apart by their entries costs one thing, and it is worth saying out loud: an entry a
/// user took out can come back, because a list trimmed to exactly the entries this app shipped before
/// that entry existed is this app's own as far as this can tell, and is read as one. What that buys
/// is the other half — the formats added since, which a list nobody had touched would otherwise
/// never be given.
///
/// The rows it walks are the ones marked repaired, and the history it reads them against is the
/// `before` of each — which is why the two are one struct rather than two tables: a list this app
/// changed and a list the repair walks are the same fact about the same list.
pub(crate) fn repair_older_lists(ini: &mut Ini) -> bool {
    let mut repaired = false;

    for list in LISTS.iter().filter(|list| list.repaired) {
        let Some(value) = ini.get(list.section, list.key) else {
            continue;
        };

        let held = list.entries.sanitize(&value);
        let canonical = list.built_in();
        let written_by_the_app = same_entries(&held, &canonical)
            || list
                .before
                .iter()
                .any(|older| same_entries(&held, &list.entries.sanitize(older)));

        // A list that already agrees with the built-in one, entry for entry and in order, is left
        // alone rather than written out again: a write moves the mtime, and the watcher would read
        // the file back for a change that was not one.
        if written_by_the_app && held != canonical {
            ini.set(list.section, list.key, Some(canonical.join(",")));
            repaired = true;
        }
    }

    repaired
}
#[cfg(test)]
mod tests {
    use super::*;

    /// The one sanitiser, read by four rules: what each of them keeps is the whole of what sixteen
    /// lists that are one function are sixteen lists.
    ///
    /// The failures are the ones the rules exist for — a compound name the archive list claims by
    /// silently ceasing to be claimed, a C# project of the text list not being a text file, and a
    /// dot file the name list cannot find because its dot was read as a separator.
    #[test]
    fn an_entry_is_read_by_the_rule_its_list_names() {
        assert_eq!(
            Entries::Bare.sanitize(".ZIP, ttf ,nonsense*,,woff2"),
            vec!["zip", "ttf", "woff2"],
            "a typed extension is what a user means, and an entry with a character in it that \
             cannot be one is dropped rather than matched against"
        );

        assert_eq!(
            Entries::Bare.sanitize("tar.gz"),
            Vec::<String>::new(),
            "a bare list has no compound name to keep, and drops the entry rather than splitting \
             one name into two extensions — which is the archive list's whole difference"
        );

        assert_eq!(
            Entries::Compound.sanitize(".ZIP, tar.gz ,nonsense*,,docx"),
            vec!["zip", "tar.gz", "docx"],
            "the archive list keeps the compound a tarball is claimed by"
        );

        assert_eq!(
            Entries::Hashed.sanitize("cs,cshtml,C#"),
            vec!["cs", "cshtml", "c#"],
            "the text list admits a hash so a C# project is a text file"
        );

        assert_eq!(
            Entries::Bare.sanitize("cs,C#"),
            vec!["cs"],
            "and no other list does, so a hash is a stray entry there"
        );

        assert_eq!(
            Entries::Name.sanitize(".gitignore,Makefile,cmakelists.txt,.env.example"),
            vec!["gitignore", "makefile", "cmakelists.txt", "env.example"],
            "a name list reads the dot as part of the name, at the front of one and inside another"
        );
    }

    /// A list is a set, so the same name typed twice is one name: two entries would be two matches
    /// where one file already satisfies the first.
    #[test]
    fn a_name_typed_twice_is_one_name() {
        assert_eq!(
            Entries::Bare.sanitize("TTF, ttf ,ttf"),
            vec!["ttf"],
            "the case an entry was typed in is not part of the name it is looked up by"
        );
    }

    /// The whole of the repair: a `config.ini` an older build wrote must come up to the list of now,
    /// for every older list this app ever shipped, and the entries must not drift.
    ///
    /// This is walked from the table rather than written out list by list, which is what it is for:
    /// an older list no row names is a format that never reaches an installation which already
    /// exists, and nothing else in the tree would say so.
    #[test]
    fn every_older_list_this_app_shipped_is_brought_up_to_the_list_of_now() {
        let mut seen = 0;

        for list in LISTS.iter().filter(|list| list.repaired) {
            let canonical = list.built_in();
            assert!(
                !canonical.is_empty(),
                "a list holds nothing to bring a file up to"
            );

            for older in list.before {
                let mut ini = Ini::new();
                ini.set(list.section, list.key, Some(older.to_string()));

                assert!(
                    repair_older_lists(&mut ini),
                    "a file holding the older `[{}]` list was not brought up",
                    list.section
                );
                assert_eq!(
                    ini.get(list.section, list.key).as_deref(),
                    Some(canonical.join(",").as_str()),
                    "a file holding the older `[{}]` list is not the list of now",
                    list.section
                );
                seen += 1;
            }
        }

        assert!(
            seen > 0,
            "no row is repaired, so no list this app changed is brought up"
        );
    }

    /// A list a person edited is theirs, and the repair is not what decides otherwise: a file
    /// holding the built-in entries and one of their own is kept exactly as it is, because the
    /// repair can only tell an edit from a list this app wrote by the entries, and one entry the
    /// app never shipped is an edit.
    #[test]
    fn a_list_a_person_edited_is_theirs_and_is_left_exactly_as_it_is() {
        for list in LISTS.iter().filter(|list| list.repaired) {
            let mut ini = Ini::new();
            let edited = format!("{},mine", list.defaults);
            ini.set(list.section, list.key, Some(edited.clone()));

            assert!(
                !repair_older_lists(&mut ini),
                "`[{}]` holds a name this app never shipped and the repair rewrote it",
                list.section
            );
            assert_eq!(
                ini.get(list.section, list.key).as_deref(),
                Some(edited.as_str())
            );
        }
    }

    /// The six lists that arrived with their kind are not walked, and this is the whole of what that
    /// buys: a file holding one of them has been edited by a person, and a person's order is theirs.
    ///
    /// It is also the one thing here that is a fact about a row rather than about a list, and the
    /// flag on the row is the only place that says so — so it is held here against a row being given
    /// the flag by accident.
    #[test]
    fn a_section_that_arrived_with_its_kind_is_not_walked() {
        for list in LISTS.iter().filter(|list| !list.repaired) {
            assert!(
                list.before.is_empty(),
                "`[{}]` is not walked and has an older list, so nothing ever reads it",
                list.section
            );
        }

        for name in ["font", "audio", "office", "ebook", "text"] {
            assert!(
                LISTS
                    .iter()
                    .any(|list| list.section == name && !list.repaired),
                "a kind that arrived with its section is not walked by the repair"
            );
        }
    }

    /// A list is written and read back as the same entries, which is what a save followed by a load
    /// has to be for the settings reset to leave a file's own lists alone.
    #[test]
    fn a_list_written_out_of_the_configuration_is_the_list_read_back_from_it() {
        let mut config = AppConfig::default();
        for list in LISTS {
            let mut list_as_held = list.built_in();
            list_as_held.push("rhn-test".to_string());
            (list.set)(&mut config, list_as_held);
        }

        let mut ini = Ini::new();
        write_all(&config, &mut ini);

        let mut read_back = AppConfig::default();
        read_all(&ini, &mut read_back);

        for list in LISTS {
            assert_eq!(
                (list.held)(&read_back),
                (list.held)(&config),
                "`[{}]` came back as something other than what was written",
                list.section
            );
        }
    }

    /// The built-in lists, as the app starts with them, are the lists a configuration that has just
    /// been made holds — which is what the lists reset promises, and what it used to reach by way of
    /// a default configuration.
    #[test]
    fn the_lists_reset_puts_the_built_in_lists_back() {
        let mut config = AppConfig::default();
        for list in LISTS {
            (list.set)(&mut config, vec!["rhn-test".to_string()]);
        }

        reset_built_in(&mut config);

        for list in LISTS {
            assert_eq!(
                (list.held)(&config),
                &list.built_in()[..],
                "`[{}]` was not put back as the built-in list",
                list.section
            );
        }
    }

    /// The settings reset has to leave the lists exactly as they are, which is the whole of what
    /// taking them out and putting them back is for.
    #[test]
    fn the_settings_reset_leaves_the_lists_exactly_as_they_are() {
        let mut config = AppConfig::default();
        for list in LISTS {
            (list.set)(&mut config, vec!["rhn-test".to_string()]);
        }
        let edited = held(&config);

        let taken = held(&config);
        let is_first_run = config.is_first_run;
        config = AppConfig::default();
        config.is_first_run = is_first_run;
        put(&mut config, taken);

        assert_eq!(
            held(&config),
            edited,
            "a reset of the settings moved a list"
        );
    }

    /// A kind added to the app is a row added to the table, and a list no row names would be a
    /// field nobody wrote: empty on a fresh configuration, with nothing anywhere to say so.
    #[test]
    fn every_row_names_a_list_a_fresh_configuration_holds() {
        let config = AppConfig::default();

        for list in LISTS {
            assert_eq!(
                (list.held)(&config),
                &list.built_in()[..],
                "`[{}]` starts as something other than its built-in list, so the row does not name \
                 the field the configuration holds it in",
                list.section
            );
        }
    }

    /// The lists are keyed by the section they are written under, and two of them may share a
    /// section only by being two keys of it — the text kind's extensions and its names, which is the
    /// only pair anywhere that does.
    #[test]
    fn a_list_is_keyed_by_its_section_and_its_key() {
        let mut keys: Vec<(&str, &str)> =
            LISTS.iter().map(|list| (list.section, list.key)).collect();
        let count = keys.len();
        keys.sort_unstable();
        keys.dedup();
        assert_eq!(
            keys.len(),
            count,
            "two rows are written under one section and key"
        );

        let shared: Vec<&str> = LISTS
            .iter()
            .map(|list| list.section)
            .collect::<std::collections::BTreeSet<_>>()
            .into_iter()
            .filter(|section| LISTS.iter().filter(|list| list.section == *section).count() > 1)
            .collect();
        assert_eq!(
            shared,
            vec!["text"],
            "one section holds two lists and it is the text one"
        );
    }
}
