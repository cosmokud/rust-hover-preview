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
//!
//! The table is in three files below: `sanitize` for what an entry of a list may be,
//! `built_ins` for the lists as this build ships them and the older ones it has shipped, and
//! `list_store` for the row itself, the sixteen of them in [`LISTS`] and the four moves over
//! them. What is left in this file is the way in — every name the rest of the tree reaches as
//! `lists::` — and the two imports the tests beside it build their file with.

#[cfg(test)]
use crate::config::config::AppConfig;
#[cfg(test)]
use configparser::ini::Ini;

mod built_ins;
mod list_store;
mod sanitize;

// The way in rather than a list of what the tree reaches: a list is read through its row and
// never through the string it holds, so a good half of what follows is named by tests alone —
// and a re-export nothing in a build reaches is an unused import to the lint that says so.
#[allow(unused_imports)]
pub use built_ins::{
    ARCHIVE_EXTENSIONS_BEFORE_THE_COMICS, CALIBRE_EXTENSIONS_BEFORE_THE_EBOOKS,
    CALIBRE_EXTENSIONS_WITH_THE_HELP_FILE, DEFAULT_ARCHIVE_EXTENSIONS, DEFAULT_AUDIO_EXTENSIONS,
    DEFAULT_CALIBRE_EXTENSIONS, DEFAULT_DESIGN_EXTENSIONS, DEFAULT_EBOOK_EXTENSIONS,
    DEFAULT_FFMPEG_EXTENSIONS, DEFAULT_FONT_EXTENSIONS, DEFAULT_IMAGE_EXTENSIONS,
    DEFAULT_LIBRE_EXTENSIONS, DEFAULT_MAGICK_EXTENSIONS, DEFAULT_OFFICE_EXTENSIONS,
    DEFAULT_PEAZIP_EXTENSIONS, DEFAULT_TEXT_EXTENSIONS, DEFAULT_TEXT_NAMES,
    DEFAULT_VECTOR_EXTENSIONS, DEFAULT_VIDEO_EXTENSIONS, DESIGN_EXTENSIONS_BEFORE_AI,
    DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE, DESIGN_EXTENSIONS_WITH_CDR,
    IMAGE_EXTENSIONS_BEFORE_AVCI, IMAGE_EXTENSIONS_BEFORE_CODEC_FORMATS,
    IMAGE_EXTENSIONS_BEFORE_DDS, IMAGE_EXTENSIONS_BEFORE_SVG, IMAGE_EXTENSIONS_WITH_SVG,
    LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ, MAGICK_EXTENSIONS_BEFORE_THE_REST,
    PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS, PEAZIP_EXTENSIONS_BEFORE_THE_EBOOKS,
    PEAZIP_EXTENSIONS_BEFORE_THE_HELP_FILE, VECTOR_EXTENSIONS_BEFORE_SVG,
    VECTOR_EXTENSIONS_BEFORE_THE_EPS_SPELLINGS, VIDEO_EXTENSIONS_BEFORE_THE_SPLIT,
};
#[allow(unused_imports)]
pub(crate) use list_store::{
    held, put, read_all, repair_older_lists, reset_built_in, write_all, List, ARCHIVE, AUDIO,
    CALIBRE, DESIGN, EBOOK, FFMPEG, FONT, IMAGE, LIBRE, LISTS, MAGICK, NAMES, OFFICE, PEAZIP, TEXT,
    VECTOR, VIDEO,
};
#[allow(unused_imports)]
pub(crate) use sanitize::Entries;
#[allow(unused_imports)]
pub use sanitize::{
    sanitize_archive_extension_list, sanitize_extension_list, sanitize_extensions, sanitize_names,
};

#[cfg(test)]
mod tests;
