//! The engine's own answer, read into the shape the readers above produce.
//!
//! Nothing here starts a tool: every function is handed a report that has already come back —
//! PeaZip's own listing of an archive, `zstd -l`'s table, zpaq's and FreeArc's lines — and reads
//! what a page of contents is made of out of it, so a hover on a format this app has no reader
//! for costs a report that was asked for once and answered by text.

use super::{normalize_name, Listing};
use std::path::Path;

/// The line the engine writes between what it says about the archive and the entries themselves.
///
/// Everything above it is the engine talking about its own run — its version, the file it
/// scanned, the archive's own header — and everything below it is the archive's table of
/// contents in the same `Key = Value` shape, one blank line between entries. Splitting there is
/// what keeps the archive's own `Path` out of the listing, since the header of the report states
/// it too.
const ENGINE_SEPARATOR: &str = "----------";

/// The listing the PeaZip engine's own report of an archive gives, read into the shape every
/// reader of this app's produces.
///
/// What the engine is asked for is a technical listing — each entry's own fields, one to a line —
/// and what it answers with is that listing as text, which is the only thing about it this side
/// has to be able to read. The fields are read for what the page needs and nothing else: the
/// entry's path, what it weighs and what it took, whether it is a folder, and whether it is
/// encrypted. Everything the engine says beyond those — the method an entry was packed with, the
/// block it shares, its CRC, its timestamps, its attributes — belongs to an extraction this app
/// never does.
///
/// Two shapes of report are worth naming, because both are ordinary and neither is an error. A
/// report with nothing under the separator is an archive the engine could not open the table of
/// contents of — one whose own headers are encrypted, which it will not list without a password —
/// and it is answered with no listing rather than with an empty one: a page saying an archive
/// holds nothing is a claim, and this is not the place that claim can be made from. And a
/// single-stream format keeps no name inside itself at all — a `bzip2` or a `zstd` file is the
/// compressed bytes and nothing beside them — so the one entry the engine reports for one of
/// those has no path, and what it is called is worked out from the archive's own name (see
/// [`stream_member_name`]).
pub(crate) fn engine_listing(report: &str, archive: &Path, file_size: u64) -> Option<Listing> {
    let (_, entries) = report.split_once(ENGINE_SEPARATOR)?;

    let mut listing = Listing::new();
    listing.file_size = file_size;

    let mut fields: Vec<(&str, &str)> = Vec::new();

    // The report ends with a blank line, and a block is ended by one — so the walk is given one
    // more line than the report has and the last block is closed by it rather than by the end of
    // the text.
    for line in entries.lines().chain(std::iter::once("")) {
        if line.trim().is_empty() {
            if !fields.is_empty() {
                push_engine_entry(&mut listing, &fields, archive);
                fields.clear();
            }
            continue;
        }

        // A field is `Name = Value`, and the value may be empty — which is how the engine says
        // that a field does not apply to this entry. The split is on the first separator, so a
        // path with one of its own is still a path.
        if let Some((name, value)) = line.split_once(" = ") {
            fields.push((name.trim(), value.trim()));
        }
    }

    // An archive with nothing in it and one whose table could not be read are the same report
    // from here, and both are answered the same way (see above).
    if listing.entries.is_empty() {
        return None;
    }

    listing.settle_totals();
    Some(listing)
}

/// One entry out of the fields the engine reported for it.
fn push_engine_entry(listing: &mut Listing, fields: &[(&str, &str)], archive: &Path) {
    let value = |name: &str| {
        fields
            .iter()
            .find(|(field, _)| *field == name)
            .map(|(_, value)| *value)
    };

    let Some(name) = value("Path").map(str::to_string).or_else(|| {
        // An entry the engine reports no path for is the member of a single-stream format, which
        // has no name inside it to report.
        value("Size")?;
        stream_member_name(archive)
    }) else {
        return;
    };

    // A folder is stated twice over by the two shapes of report: the entry's own attributes
    // carry the `D` every extractor reads, and the containers that keep a separate flag state it
    // as one of the entry's fields.
    let is_dir = value("Attributes").is_some_and(|attributes| attributes.contains('D'))
        || value("Folder") == Some("+");

    let size = value("Size")
        .and_then(|size| size.parse::<u64>().ok())
        .unwrap_or(0);
    // A field left empty is a field the engine has no answer for — a member of a solid block
    // shares its packed size with the entries beside it — and a total the page cannot state is
    // what `settle_totals` does with one.
    let packed = value("Packed Size").and_then(|packed| packed.parse::<u64>().ok());
    let encrypted = value("Encrypted") == Some("+");

    if let Some(name) = normalize_name(&name) {
        listing.push(
            name,
            if is_dir { 0 } else { size },
            packed,
            is_dir,
            encrypted,
        );
    }
}

/// The listing of a single-stream file whose own tool has nothing to say: the member an extraction
/// would write.
///
/// This is the one answer here that no tool produced. Brotli, BCM and LPAQ put one file into one
/// file and print nothing about what is inside — there is nothing inside but the bytes — so there
/// is nothing to ask, and what a hover on one of their files shows is what is true without asking:
/// that there is one member, what an extraction would call it, and how much the file it came from
/// weighs. What is *not* stated is the member's own size: none of those three streams keeps one,
/// and the only way to learn it is to decompress the whole file, which is not a hover's work. The
/// page says what it knows and no more — the name, and the file's weight on disk.
pub(crate) fn stream_listing(path: &Path, file_size: u64) -> Option<Listing> {
    let name = normalize_name(&stream_member_name(path)?)?;

    let mut listing = Listing::new();
    listing.file_size = file_size;
    listing.push(name, 0, None, false, false);
    listing.settle_totals();

    Some(listing)
}

/// The listing `zstd -l` gives for a `.zst`, read into the shape every reader here produces.
///
/// What the tool prints is a header and one line for the file it was given: how many frames the
/// stream holds, how much it weighs, how much it weighed before it was compressed, and where it
/// is. The member is the one the stream is — a `.zst` keeps no name inside itself any more than a
/// `bzip2` does, so it is named the way every single-stream member here is (see
/// [`stream_member_name`]) — and what the tool adds is the size, which is what this route is for:
/// the console archiver lists a `.zst` too and leaves that column blank.
///
/// A file zstd will not read is answered with the header and a line of its own saying so rather
/// than with a data line, and that is no listing — the same answer an archive the archiver cannot
/// open gets from it.
pub(crate) fn zstd_listing(report: &str, archive: &Path, file_size: u64) -> Option<Listing> {
    let name = normalize_name(&stream_member_name(archive)?)?;

    // The data line is the one whose first two columns are counts: the header's are words, and a
    // complaint's first word is not a number.
    let line = report.lines().find(|line| {
        let mut fields = line.split_whitespace();
        matches!((fields.next(), fields.next()),
            (Some(frames), Some(skips))
                if frames.parse::<u64>().is_ok() && skips.parse::<u64>().is_ok())
    })?;

    let mut fields = line.split_whitespace().skip(2).peekable();
    let packed = stream_size(&mut fields)?;
    // A stream whose frame header states no size is one the tool cannot weigh, and a member of
    // unknown size is shown as nothing rather than as a number this side does not have.
    let size = stream_size(&mut fields).unwrap_or(0);

    let mut listing = Listing::new();
    listing.file_size = file_size;
    listing.push(name, size, Some(packed), false, false);
    listing.settle_totals();

    Some(listing)
}

/// The listing zpaq gives for a `.zpaq`, read into the shape every reader here produces.
///
/// zpaq is a journaling archiver, and what it prints for a name is the latest version of it — one
/// line for a file, with the date and time it was saved, how large it is, and a flag before the
/// name: `#` for a folder and `+` for a file. A name is the last thing on the line and may hold
/// spaces, so it is taken as what is left of the line after the flag rather than as a column, and
/// only lines that open with a date and a time are read at all: the header above them, the totals
/// below them and the tool's own closing line are not entries.
///
/// Nothing else of a line is kept. What zpaq stores of a file is a date, its attributes and the
/// fragments its contents were split into, and none of that is what a page of contents is made of
/// — the sizes are the whole of what a hover shows.
pub(crate) fn zpaq_listing(report: &str, file_size: u64) -> Option<Listing> {
    let mut listing = Listing::new();
    listing.file_size = file_size;

    for line in report.lines() {
        let words = words_of(line);
        if !opens_with_a_timestamp(&words) {
            continue;
        }

        // What follows the timestamp: the size, the ratio, the flag and then the name. A folder is
        // marked by a bracket before its size, which is the one line whose fields do not start
        // where every other line's do.
        let fields = if words.get(2).is_some_and(|(_, word)| *word == "[") {
            &words[3..]
        } else {
            &words[2..]
        };

        let (Some((_, size)), Some((_, ratio)), Some((_, flag)), Some((name_start, _))) =
            (fields.first(), fields.get(1), fields.get(2), fields.get(3))
        else {
            continue;
        };

        // The ratio column is a percentage or the word a folder is marked with, and the flag is
        // one of the two characters the tool marks a line with. A line that is neither is a line
        // of another shape, and it is skipped rather than guessed at.
        let Ok(size) = size.parse::<u64>() else {
            continue;
        };
        if !(ratio.ends_with('%') || *ratio == "dir") || !matches!(*flag, "+" | "#") {
            continue;
        }

        let is_dir = *flag == "#";
        if let Some(name) = normalize_name(line[*name_start..].trim()) {
            listing.push(name, if is_dir { 0 } else { size }, None, is_dir, false);
        }
    }

    if listing.entries.is_empty() {
        return None;
    }

    listing.settle_totals();
    Some(listing)
}

/// The listing FreeArc's verbose report gives for an `.arc`, read into the shape every reader here
/// produces.
///
/// What it is asked for is `v` — its verbose listing — and what it prints is a line per entry with
/// the date, the attributes, the size, the space the entry takes and its CRC, and then the name.
/// It is the one of FreeArc's three listings that carries the attributes, which is where a folder
/// is told from a file: a directory's are marked with a `D`, a file's are dots.
///
/// The packed column is deliberately not read. What FreeArc reports there is block accounting
/// rather than a size per entry — the second file of a solid block is reported as taking nothing,
/// because the block it shares was paid for by the first — so a sum of that column is not the
/// archive's compressed size, and a page built on one would state a saving that is not a saving.
/// What the page gives instead is what the file itself weighs on disk.
pub(crate) fn arc_listing(report: &str, file_size: u64) -> Option<Listing> {
    let mut listing = Listing::new();
    listing.file_size = file_size;

    for line in report.lines() {
        let words = words_of(line);
        if !opens_with_a_timestamp(&words) {
            continue;
        }

        // Attributes, the size, the space taken and the CRC — that is four columns — and then the
        // name, which is the last of them and may hold spaces, so it is what is left of the line
        // after the CRC rather than a column of its own.
        let (Some((_, attributes)), Some((_, size)), Some(_), Some((crc_start, crc))) =
            (words.get(2), words.get(3), words.get(4), words.get(5))
        else {
            continue;
        };

        let Ok(size) = size.parse::<u64>() else {
            continue;
        };

        let is_dir = attributes.contains('D');
        let name = line[crc_start + crc.len()..].trim();

        if let Some(name) = normalize_name(name) {
            listing.push(name, if is_dir { 0 } else { size }, None, is_dir, false);
        }
    }

    if listing.entries.is_empty() {
        return None;
    }

    listing.settle_totals();
    Some(listing)
}

/// The words on a line, each with the index it starts at, which is what a name that may hold
/// spaces is taken as the rest of the line from.
fn words_of(line: &str) -> Vec<(usize, &str)> {
    let mut words: Vec<(usize, &str)> = Vec::new();
    let mut start: Option<usize> = None;

    for (index, character) in line.char_indices() {
        if character.is_whitespace() {
            if let Some(start) = start.take() {
                words.push((start, &line[start..index]));
            }
        } else if start.is_none() {
            start = Some(index);
        }
    }

    if let Some(start) = start {
        words.push((start, &line[start..]));
    }

    words
}

/// Whether a line opens with the date and the time both FreeArc and zpaq put in front of an entry
/// — `2026-09-25` and `12:15:27` — which is what an entry line is told by, since nothing else in
/// either report opens with one.
fn opens_with_a_timestamp(words: &[(usize, &str)]) -> bool {
    let (Some((_, date)), Some((_, time))) = (words.first(), words.get(1)) else {
        return false;
    };

    let at =
        |value: &str, index: usize, separator: u8| value.as_bytes().get(index) == Some(&separator);

    date.len() == 10
        && at(date, 4, b'-')
        && at(date, 7, b'-')
        && time.len() == 8
        && at(time, 2, b':')
        && at(time, 5, b':')
}

/// One size out of `zstd -l`'s table: a number and, where the tool wrote one, the unit it is in.
fn stream_size<'a>(fields: &mut std::iter::Peekable<impl Iterator<Item = &'a str>>) -> Option<u64> {
    let value: f64 = fields.next()?.parse().ok()?;

    let multiplier = match fields.peek().and_then(|unit| unit_in_bytes(unit)) {
        Some(multiplier) => {
            fields.next();
            multiplier
        }
        None => 1.0,
    };

    Some((value * multiplier).round() as u64)
}

/// What one of the units `zstd -l` writes after a size is worth in bytes, and nothing for a token
/// that is not one — that token is the next column's and is left where it is.
fn unit_in_bytes(unit: &str) -> Option<f64> {
    match unit {
        "B" => Some(1.0),
        "KiB" => Some(1024.0),
        "MiB" => Some(1024.0 * 1024.0),
        "GiB" => Some(1024.0 * 1024.0 * 1024.0),
        "TiB" => Some(1024.0 * 1024.0 * 1024.0 * 1024.0),
        _ => None,
    }
}

/// What the one member of a single-stream format is called: the archive's own name with the
/// extension it was compressed under taken off, which is the name an extraction writes it under
/// and the name the file had before it was compressed.
///
/// Nothing here can recover the name it *did* have — a `bzip2` stream keeps no name, and a `zstd`
/// one usually keeps none either — so what is shown is the closest true thing rather than a guess
/// at the original.
fn stream_member_name(archive: &Path) -> Option<String> {
    let name = archive.file_name().and_then(|name| name.to_str())?;

    let stem = name
        .rsplit_once('.')
        .map(|(stem, _)| stem)
        .filter(|stem| !stem.is_empty())
        .unwrap_or(name);

    Some(stem.to_string())
}
