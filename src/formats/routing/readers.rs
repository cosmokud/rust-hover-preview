//! Which readers could answer, resolved to this machine.
//!
//! Split out of [`super`] because it is the one question in this module that is about the machine
//! and the run rather than about the file: the order `chain` declares belongs to every machine,
//! and whether a reader in it is installed is this machine's own answer, asked once per engine
//! per preview. The two are kept apart for that reason, and the parent re-exports what is public,
//! so every caller still reaches it as `routing::readers_for`.

use super::chain;
use crate::config::config::PreviewType;
use crate::engines::{
    calibre_render, imagemagick_render, libreoffice_render, peazip_render, webview_preview,
};
use crate::formats::{codecs, office_formats};
use std::path::Path;

/// The readers of a file of this kind that could answer for it, in the order they are asked.
///
/// It is [`chain`] narrowed to what is true of the machine and the file: a reader that is not
/// installed is not one to ask, one that was already asked about this file and turned it down
/// is not asked twice, and one that cannot be started here is not a wait to sit through. What
/// is left is asked in the chain's order, so the first reader that can answer is the one that
/// does — which is what the preview loop walks when it has a hover to fill and no page to fill
/// it with.
///
/// Whether a reader has *already* answered is not asked here: a page that has landed is in
/// hand, and what the loop asks of this is only which engine to ask next. That question is
/// asked where the answers are kept (see the render tiers in `preview_window`).
pub fn readers_for(kind: PreviewType, path: &Path) -> Vec<Reader> {
    chain(kind)
        .iter()
        .copied()
        .filter(|reader| can_answer(*reader, path))
        .collect()
}

/// Whether a reader is one to ask about `path` at all.
fn can_answer(reader: Reader, path: &Path) -> bool {
    match reader {
        // A reader of this app's own is always there, and what it can read is the file's own
        // question rather than the machine's (`native_formats`).
        Reader::Native => true,

        // The application that owns the format, where the machine has it: the answer is per
        // file rather than per engine, because which application a name belongs to is the
        // name's (see `office_formats::app_for`).
        Reader::Office => office_formats::app_installed(path),

        Reader::LibreOffice => {
            libreoffice_render::available() && !libreoffice_render::refused(path)
        }
        Reader::ImageMagick => {
            imagemagick_render::available() && !imagemagick_render::refused(path)
        }
        // Per name rather than per engine: the tools beside the console archiver read three of
        // the names in the list, and a machine with the archiver and without the tool is a
        // machine that cannot list those three (see `peazip_formats::Backend`).
        Reader::PeaZip => peazip_render::available_for(path) && !peazip_render::refused(path),
        Reader::Calibre => calibre_render::available() && !calibre_render::refused(path),

        // FFmpeg's player, where it is installed: the one reader preferred over the engine
        // Windows has, and the reason a video is a chain rather than a reader (see `codecs`).
        Reader::Ffmpeg => codecs::ffplay_available(),

        // The browser engine, which draws what neither of the two others can: a metafile is
        // not this reader's at all, but a document drawn by it is nothing if it is missing.
        Reader::WebView2 => webview_preview::is_available(),
    }
}

/// One reader of a file: a reader of this app's own, or an engine this app drives.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Reader {
    /// A reader of this app's own — the picture decoders, the text and archive readers, the
    /// PDF engine Windows has, the media engine that plays a video. Which one is a question
    /// of the file, and it is asked by `native_formats`.
    Native,
    /// The application that owns the format, driven through its own automation.
    Office,
    /// The installed render engine, which draws a page of what its import filters read.
    LibreOffice,
    /// The installed image converter, which develops a picture nothing else opens.
    ImageMagick,
    /// The installed archiver, which lists a container no reader here opens.
    PeaZip,
    /// The installed ebook converter, which writes a PDF of a book nothing here reads.
    Calibre,
    /// FFmpeg's player, where it is installed: the engine video previews are played by when
    /// it is, and the one engine preferred over the reader Windows has.
    Ffmpeg,
    /// The browser engine the machine already has, which draws SVG documents and font
    /// specimens — neither of them rasterized on this side.
    WebView2,
}
