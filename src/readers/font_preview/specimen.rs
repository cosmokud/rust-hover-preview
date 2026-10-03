use super::font_tables::{
    collection_face, extract_face, file_stamp, name_strings, path_hash, tables_of, CharacterMap,
};
use crate::config::config::{read_within_budget, DEFAULT_TTC_FACE};
use crate::CONFIG;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;

/// The box a font preview is measured at, in the same units a document's own size is:
/// what the layout places the engine's window by.
///
/// A font has no size it asks to be drawn at — what a file holds is outlines, and the text
/// they are drawn as is whatever size a caller asks for — so the box a specimen is placed
/// in is this app's own: a page a shade wider than tall, which is the shape the pangram and
/// a line or two under it want. What the setting beside it names is the share of the
/// display that box takes, the way `vector_scale` names one for a drawing.
pub const SPECIMEN_WIDTH: u32 = 1500;
pub const SPECIMEN_HEIGHT: u32 = 1000;

/// The lines a specimen may be drawn from, in the order they are listed: the pangram every
/// Latin font is judged by, and then one line for each script a font may hold instead of —
/// or as well as — Latin.
///
/// A line is drawn only where the font's own character map covers every character of it;
/// see this module's documentation. Each is the closest thing its script has to a pangram,
/// which is what a specimen is for: the Japanese line is the iroha, the poem the kana were
/// ordered by for a thousand years; the Chinese one is the first line of the Thousand
/// Character Classic, four characters of which carry almost every stroke a Han glyph is
/// built from; the Cyrillic, Greek, Arabic and Hebrew ones are the pangrams those scripts
/// are sampled with; and the Thai and Devanagari ones are the openings of the pangrams
/// those scripts are sampled with, the rest of each being words a font is as likely to be
/// asked for as it is to have — and a line is drawn whole or not at all, so a word a font
/// has not got costs the script its line rather than part of one.
pub(super) const SAMPLE_LINES: [&str; 10] = [
    "The quick brown fox jumps over the lazy dog.",
    "いろはにほへと ちりぬるを",
    "天地玄黄 宇宙洪荒",
    "다람쥐 헌 쳇바퀴에 타고파",
    "Съешь же ещё этих мягких французских булок да выпей чаю",
    "Ξεσκεπάζω την ψυχοφθόρα βδελυγμία",
    "نص حكيم له سر قاطع وذو شأن عظيم مكتوب على ثوب أخضر ومغلف بجلد أزرق.",
    "דג סקרן שט בים מאוכזב ולפתע מצא חברה",
    "เป็นมนุษย์สุดประเสริฐเลิศคุณค่า กว่าบรรดาฝูงสัตว์เดรัจฉาน",
    "ऋषियों को सताने वाले दुष्ट राक्षसों के राजा रावण का सर्वनाश",
];

/// How many characters the last-resort line carries: what a font that covers none of the
/// lines above is shown by, taken from its own character map. A line of two dozen glyphs is
/// enough to see what a face looks like, and short enough that the line stays one.
const FALLBACK_LINE_CHARACTERS: usize = 32;

/// A specimen: the title the preview is headed with, and the lines drawn under it.
///
/// The first line is the headline — the pangram wherever the font has Latin, and the
/// font's own characters where it has none of the lines this app knows — and the rest are
/// drawn smaller, one script apiece.
#[derive(Clone, PartialEq, Eq, Debug)]
pub struct Specimen {
    pub title: String,
    pub samples: Vec<String>,
    /// How many faces the file holds: more than one is a collection, which the engine can
    /// be pointed at only through a face written out on its own; see `browser_source`.
    pub faces: usize,
    /// Which of those faces this is, as an index into the collection's directory — the face
    /// the setting asked for, reduced to the last one the file holds where it holds fewer
    /// than that, and `0` for a file that is not a collection at all. It is the face the
    /// title counts from one and the one `browser_source` writes out.
    pub face: usize,
}

/// The specimen `path` makes at the face the setting names: see `configured_face`.
///
/// The answer is held between hovers — the file, the version of it that was read, and the
/// face — for the reason a document's measurement is: the layout asks for it on every
/// hover, and a pointer swept back and forth over a folder meets the same files again. A
/// file that turned out not to be a font at all is held as that, so a hover onto it costs
/// nothing after the first.
pub fn probe(path: &Path) -> Option<Specimen> {
    probe_face(path, configured_face())
}

/// The specimen `path` makes at one face of a collection: `face` is an index into the
/// faces a `.ttc` holds, and a file that holds a single font answers it the same way
/// whatever it is.
///
/// The face is part of what a held specimen is valid for, because it is part of what the
/// specimen is: the two faces of one collection are two fonts that share a file, and an
/// answer read at one of them is not an answer for the other.
pub fn probe_face(path: &Path, face: usize) -> Option<Specimen> {
    let key = FontKey {
        path: path.to_path_buf(),
        version: file_version(path),
        face,
    };

    if let Some(held) = held(&key) {
        return held;
    }

    let specimen = read_specimen(path, face);
    hold(&key, specimen.clone());

    specimen
}

/// Which face of a collection a preview is of, as the setting numbers it — from `1`, the
/// first face, down to the last face a file holds.
///
/// It is read from the configuration each time rather than captured, so the tray's `Font
/// Face` setting applies to the next hover rather than to the next run. What comes back is
/// the index the face is read at, which is the setting's own number less one; a file with
/// fewer faces than that is read at the last one it has, so a collection of two is its
/// second face whatever past the second the setting asks for.
pub fn configured_face() -> usize {
    CONFIG
        .lock()
        .map(|config| config.ttc_face)
        .unwrap_or(DEFAULT_TTC_FACE)
        .saturating_sub(1) as usize
}

/// The file the engine is pointed at for `path`: the font itself, or — for a collection,
/// which no page can name a face of — the face the specimen was read at, written out as a
/// font of its own.
///
/// The extracted face lands in `folder`, which is the browser's own profile folder for this
/// run; it is named for the file, the version of it that was read and the face, so the same
/// face of the same collection hovered again is the same file, another face is another, and
/// an edited one is another again. What this costs is one write per face per version, and
/// what it is not is permanent: the folder is cleared at the next start, with the rest of
/// what a run leaves behind.
pub fn browser_source(path: &Path, specimen: &Specimen, folder: &Path) -> Option<PathBuf> {
    if specimen.faces <= 1 {
        return Some(path.to_path_buf());
    }

    let (offset, flavor) = collection_face(path, specimen.face)?;
    let extension = if flavor == *b"OTTO" { "otf" } else { "ttf" };
    let extracted = folder.join("fonts").join(format!(
        "{:016x}-{}-{}.{extension}",
        path_hash(path),
        file_stamp(path),
        specimen.face
    ));

    if extracted.is_file() {
        return Some(extracted);
    }

    let bytes = read_within_budget(path)?;
    let face = extract_face(&bytes, offset)?;

    std::fs::create_dir_all(extracted.parent()?).ok()?;
    std::fs::write(&extracted, face).ok()?;

    Some(extracted)
}

/// The specimen a file makes at one of its faces, read fresh.
fn read_specimen(path: &Path, face: usize) -> Option<Specimen> {
    let bytes = read_within_budget(path)?;
    let tables = tables_of(&bytes, face)?;

    // What a specimen is checked against, and what it is drawn from where none of the
    // lines fit: a font with no character map has no characters, so there is nothing to
    // show and nothing to check.
    let cmap = tables.cmap.as_deref().and_then(CharacterMap::parse)?;
    let samples = sample_lines(&cmap);

    if samples.is_empty() {
        return None;
    }

    Some(Specimen {
        title: specimen_title(path, tables.name.as_deref(), tables.faces, tables.face),
        samples,
        faces: tables.faces,
        face: tables.face,
    })
}

/// The lines a specimen is drawn from: every line of the ten the font's own character map
/// covers, or — where it covers none of them — one line of its own characters.
fn sample_lines(cmap: &CharacterMap) -> Vec<String> {
    let covered: Vec<String> = SAMPLE_LINES
        .iter()
        .filter(|line| covers(cmap, line))
        .map(|line| line.to_string())
        .collect();

    if !covered.is_empty() {
        return covered;
    }

    let characters = cmap.drawable_codes(FALLBACK_LINE_CHARACTERS);

    if characters.is_empty() {
        Vec::new()
    } else {
        vec![characters.into_iter().collect()]
    }
}

/// Whether a font can draw every character of a line.
///
/// All or nothing, and the space between two words is not a character a font has to have:
/// a line drawn with one character falling back to a system font would be a preview of two
/// fonts at once, which is exactly what the coverage check exists to keep off the screen.
fn covers(cmap: &CharacterMap, line: &str) -> bool {
    line.chars()
        .filter(|character| !character.is_whitespace())
        .all(|character| cmap.maps(character))
}

/// What a specimen is headed with: the family and style the font calls itself, or the
/// file's own name where its `name` table is missing or says nothing.
///
/// A collection is noted as such and noted for which of its faces is drawn, because what
/// is drawn is one face of it: the file is a font, and the preview is of the face the
/// setting asked for — or of the last one the file holds, where it holds fewer than that.
fn specimen_title(path: &Path, name: Option<&[u8]>, faces: usize, face: usize) -> String {
    let named = name
        .and_then(name_strings)
        .map(|(family, style)| match style {
            Some(style) if !style.is_empty() => format!("{family} {style}"),
            _ => family,
        });

    let title = named.unwrap_or_else(|| {
        path.file_stem()
            .or_else(|| path.file_name())
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default()
    });

    if faces > 1 {
        format!("{title} ({} of {faces})", face + 1)
    } else {
        title
    }
}

/// How many specimens are held at once.
///
/// What a pointer meets again is the folder it is in, so the handful of fonts under it are
/// the ones worth keeping; one that falls out is read again the next time it is hovered,
/// which is the cost this cache exists to save rather than one it turns into a failure.
const MAX_HELD_SPECIMENS: usize = 32;

/// The file's modification time and length: what says a file is not the one that was read
/// last time.
#[derive(Clone, PartialEq, Eq, Hash)]
struct FileVersion {
    modified: Option<SystemTime>,
    len: u64,
}

/// What a held specimen is valid for: the file, the version of it that was read, and the
/// face of a collection it was read at — since another face of the same file is another
/// specimen, and a setting that names one is answered only by the specimen read at it.
#[derive(Clone, PartialEq, Eq, Hash)]
struct FontKey {
    path: PathBuf,
    version: FileVersion,
    face: usize,
}

/// A held specimen and when it was last asked for. The stamp is a counter rather than a
/// clock, so the order specimens are dropped in cannot be changed by the system clock
/// moving.
struct HeldSpecimen {
    specimen: Option<Specimen>,
    last_used: u64,
}

#[derive(Default)]
struct SpecimenCache {
    entries: HashMap<FontKey, HeldSpecimen>,
    tick: u64,
}

/// The specimens held between hovers. A file that turned out not to be a font is held as
/// that, rather than as nothing held at all, so a hover onto it costs nothing after the
/// first — the same way a document that will not parse is remembered as one.
static SPECIMENS: Lazy<Mutex<SpecimenCache>> = Lazy::new(|| Mutex::new(SpecimenCache::default()));

/// What is held for a version of a file, when anything is.
fn held(key: &FontKey) -> Option<Option<Specimen>> {
    let mut cache = SPECIMENS.lock().ok()?;
    cache.tick += 1;
    let tick = cache.tick;

    let held = cache.entries.get_mut(key)?;
    held.last_used = tick;

    Some(held.specimen.clone())
}

/// Hold what a file turned out to be, dropping the least recently used specimen once the
/// cap is passed.
fn hold(key: &FontKey, specimen: Option<Specimen>) {
    let Ok(mut cache) = SPECIMENS.lock() else {
        return;
    };

    cache.tick += 1;
    let tick = cache.tick;

    cache.entries.insert(
        key.clone(),
        HeldSpecimen {
            specimen,
            last_used: tick,
        },
    );

    while cache.entries.len() > MAX_HELD_SPECIMENS {
        let Some(oldest) = cache
            .entries
            .iter()
            .min_by_key(|(_, held)| held.last_used)
            .map(|(key, _)| key.clone())
        else {
            break;
        };

        cache.entries.remove(&oldest);
    }
}

fn file_version(path: &Path) -> FileVersion {
    match std::fs::metadata(path) {
        Ok(metadata) => FileVersion {
            modified: metadata.modified().ok(),
            len: metadata.len(),
        },
        Err(_) => FileVersion {
            modified: None,
            len: 0,
        },
    }
}
