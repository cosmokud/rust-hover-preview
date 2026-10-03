//! What a file's own bytes say it is, where that is not what its name says.
//!
//! One engine — the application Word, Excel and PowerPoint are started through, the
//! LibreOffice a document is converted by, the player FFmpeg draws a video in — is only
//! ever handed a file whose content belongs to it. The Office tier has asked its own
//! header what it is before a path was given to it for as long as it has existed (see
//! `office_formats::container_kind`); what this module adds is the same question asked
//! before *every* kind, and an answer that is finer than yes or no: a `.docx` whose bytes
//! are an MP4 is not merely turned down, it is played by the player a video is played by,
//! and a `.cdr` that is really a PNG gets the picture preview. A file whose content is a
//! format no kind of this app previews — an executable, an audio file, a format nothing
//! here reads — is answered with no preview, which is the answer that leaves every engine
//! unstarted.
//!
//! Three tables answer what a file is, and what is in each of them is what it is for.
//! The common formats — the pictures, the containers a video arrives in, a PDF, the fonts,
//! the archives — are the [`infer`] crate's: a small table of the signatures every tool
//! agrees on, no dependencies of its own, and the same answer any file manager would give.
//! The formats it does not carry but a list of this app's does are this app's own table,
//! asked in the shape the rest of the app already asks signatures in: a transport stream
//! by its sync byte, an MXF by its partition pack, a raw H.264 stream by its sequence
//! parameter set, a WordPerfect document by the signature every file of the family opens
//! with — the same kind of question `office_formats::container_kind` asks an Office
//! document and `has_pdf_header_in` asks a PDF, and one of them, the transport stream, is
//! asked of `video_formats` rather than written again.
//!
//! That second table is the whole of what the two engines' lists hold and the common one
//! does not: every name `video_formats` carries that FFmpeg demuxes, and every name
//! `libre_formats` carries that the render engine imports, less the names [`infer`] has
//! already answered for. What a test is worth differs from format to format and is said
//! with each entry: a magic where the format has one, a start code and a header where the
//! format is a raw stream, a shape where the format's own demuxer scores one instead of
//! reading a magic (see `signatures::SIGNATURES` for what is in it, where each answer comes from, and
//! for the three kinds of name that are deliberately not in it).
//!
//! What no table answers is answered by the name the file carries, which is the third
//! table: the few names a head cannot settle at all — a container that says nothing about
//! itself inside the probe, a format whose signature is its own text, a name no demuxer
//! reads — each of which keeps the answer it always had. See `probe::KIND_BY_NAME` for what is
//! in it and why each entry is there.
//!
//! A table of names is a weaker answer than a table of signatures, and it is asked where
//! the strong ones have nothing to say. What it can answer is only what a name says, so a
//! file that carries one of its names and is not the format that name is written for would
//! be routed by it wrongly — which is why a name whose format is *two* formats, with the
//! engine reading one of them and nothing here previewing the other, is asked a question
//! of its own before it is answered: `.pdb` is a Palm OS database, which the render
//! engine's filters read as an ebook, and it is also the Microsoft program database a
//! compiler writes beside its binaries, which is no document at all — see
//! `palm_ebook_or_program_database` for what is asked of the bytes there. **An engine is
//! handed a file of such a name only where its own header says it is the format that
//! engine reads**, and a file that is the other one shows nothing rather than starting an
//! engine it is not for.
//!
//! What is deliberately in none of the tables matters as much as what is. **A container
//! is not a kind**, so nothing here answers with a box: a zip is what an OpenDocument, an
//! iWork document and an Office package all arrive in, the OLE compound file is what every
//! `.doc`, `.xls` and `.ppt` is, and the file's own name is a better answer than the box
//! is — which is also why a `.docx` that is really a `.doc` is still handed to an engine.
//! The one container that *is* answered for is the package that declares its own type at
//! an offset every file of the format agrees on — an OpenDocument, a StarOffice XML
//! document, a Krita project — because what is read there is the document's own answer
//! rather than the box's (see `signatures::Matcher::Package`). A format whose signature *is* its
//! text — an SVG, an HTML page, a JSON file, a flat OpenDocument — is left alone for the
//! same reason: what it is is what the text lists decide, and a name another kind already
//! read is not improved by a prefix match.
//!
//! The probe is one read of [`crate::formats::head::PROBE_BYTES`] and every question asked here is about the
//! head of a file. A file whose content is not on this machine is not opened at all, which
//! is the rule every other read in this app follows (see `cloud_files`).
//!
//! What the tables answer is read as one of three things:
//!
//! * A kind, where the format is one this app's lists claim and it is not the kind the
//!   name claimed: that kind is what the file is previewed as. The third table answers a
//!   kind outright, because there is nothing for a name to disagree with — the name is the
//!   question — which is what makes a format with no signature worth writing down at all.
//! * [`Content::Foreign`], where the format is one no list claims — an executable, an
//!   audio file, a box of some kind this app has no preview for — or where the third
//!   table's name turned out to be the other format's: a `.pdb` that is not the ebook the
//!   engine reads is a file with nothing to show, and no engine is started for it.
//! * [`Content::Unknown`], which is "no opinion" and is the answer in three cases: nothing
//!   in the tables matched and the name is not one of the third table's either (a raw
//!   stream, an exotic container, a text file), the format is one the file's own name
//!   already names (the two agree, so there is nothing to override), and a file with no
//!   extension at all, whose name has nothing to disagree with. The name decides, as it
//!   always has.
//!
//! The answer is held between hovers, keyed by the file and the version of it that was
//! read, because one hover asks this three times — the hook that raises it, the loader
//! that fills it, and the layout that places it — and a file edited in place is read again
//! rather than held to an answer about what it used to be.

mod matchers;
mod probe;
mod signatures;

pub(crate) use probe::of_reaching_config;
pub use probe::{answer, of_with_facts, Content, Probe};

// What the tests ask this module by: the question a caller asks rather than the one the
// engines ask, and the names behind it they read the tables through.
#[cfg(test)]
use crate::config::config::{AppConfig, PreviewType};
#[cfg(test)]
use matchers::is_declared_package;
#[cfg(test)]
use probe::{classify, detected_names, kind_by_name, KIND_BY_NAME};
#[cfg(test)]
pub(crate) use probe::{count_entry_reads_from_now, entry_reads, of};
#[cfg(test)]
use signatures::SIGNATURES;
#[cfg(test)]
use std::path::{Path, PathBuf};

#[cfg(test)]
mod tests;
