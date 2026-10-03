//! Font previews: what a font is, the lines its own character map covers, and the file the
//! browser engine draws it from.
//!
//! A font is not a picture to a decoder either, and this app does not rasterize one — the
//! release binary carries no rasterizer, no shaping stack and no font parser of the kind
//! that draws text. What draws a font is the same browser engine the SVG previews use, in
//! a window of its own, pointed at a page of this app's own with the font in it through
//! `@font-face`; see `webview_preview`. That engine reads all five formats a font goes by,
//! TrueType and CFF outlines, the two webfont containers and a collection of faces, so the
//! drawing costs this app no reader at all.
//!
//! What is left here is the specimen, and the one question the engine cannot answer.
//!
//! The question is what a font *covers*. A browser falls back per glyph and says nothing
//! about it: a page that draws 「いろは」 in a Latin-only font draws it in a system font, at
//! the same size and in the same layout, and a preview that showed it would be claiming
//! something about the font that is not true. So the sample lines *are* the font's own
//! coverage: each one is checked against the font's character map, and only the lines the
//! map answers for are drawn — the pangram always among them where the font has Latin, and
//! a line apiece where it has Japanese, Chinese, Korean, Cyrillic, Greek, Arabic, Hebrew,
//! Thai or Devanagari. A font of a script there is no line for — a Georgian one, a symbol
//! one — is drawn from the characters its own map holds, and so is a font with no script at
//! all: an icon font's glyphs are private-use ones, and a line of them is what a specimen of
//! such a font is for. Both are the same answer reached the other way round.
//!
//! Answering it costs a read of the file under the budget every other read is answered
//! under, and a parse of two of its tables: the character map, and the `name` table the
//! preview is titled with. Nothing is rasterized, nothing is held decoded, and what is kept
//! between hovers is the answer — keyed by the file and the version of it that was read, the
//! way an SVG document's measurement is. The two table readers are a WOFF2 font's Brotli
//! stream and a WOFF one's per-table zlib, which is all that stands between a webfont and
//! the same two tables a `.ttf` holds in the open.
//!
//! One thing is written to disk, and it is the one thing the engine cannot be handed: a page
//! has no syntax for naming a face inside a collection, so a `.ttc` is answered with the face
//! the setting names written out as a font of its own, beside the browser's own profile
//! folder and named for the file, the version of it and the face — where the next run's
//! startup clears it away with everything else that folder held.

mod font_tables;
mod specimen;

pub use specimen::*;

#[cfg(test)]
use font_tables::{tables_of, CharacterMap, WOFF2_KNOWN_TAGS};
#[cfg(test)]
use specimen::SAMPLE_LINES;
#[cfg(test)]
use std::path::PathBuf;

#[cfg(test)]
mod tests;
