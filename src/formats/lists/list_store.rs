//! The row that is a list: what each of the sixteen holds, where the file writes it, and the four
//! moves over the table — the read out of a file, the write into one, the two resets and the
//! repair that brings a file an older build wrote up to the list of now.
//!
//! The entries a row is read by are named rather than inlined, so a change of rule is a change of
//! one [`Entries`] and not of sixteen call sites; the lists themselves are in `built_ins`.

use super::built_ins::{
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
use super::sanitize::Entries;
use crate::config::config::AppConfig;
use configparser::ini::Ini;
use std::path::Path;

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
