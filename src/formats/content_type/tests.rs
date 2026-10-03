use super::*;

/// A file of one name holding the front of a format, answered the way a hover asks it:
/// the bytes through the two signature tables, and the name through the third.
fn classified(name: &str, content: &[u8]) -> Content {
    if let Some(names) = detected_names(content) {
        return classify(Path::new(name), names, &AppConfig::default());
    }

    crate::formats::text_formats::lookup_extension(Path::new(name))
        .and_then(|extension| kind_by_name(&extension, content))
        .unwrap_or(Content::Unknown)
}

/// The database header a Palm OS document opens with: the name it is filed under, the
/// four-character type and creator of the application that wrote it, and the count of
/// the records that follow.
fn palm_database(name: &str, kind: &[u8; 4], creator: &[u8; 4]) -> Vec<u8> {
    let mut header = vec![0u8; 78 + 8];
    header[..name.len()].copy_from_slice(name.as_bytes());
    header[60..64].copy_from_slice(kind);
    header[64..68].copy_from_slice(creator);
    header[76..78].copy_from_slice(&1u16.to_be_bytes());
    header
}

/// The same header with the two identifiers a Mobipocket book is defined by — the type `BOOK`
/// and the creator `MOBI` — and the record a MOBI header of its own begins in.
fn mobipocket() -> Vec<u8> {
    let mut header = palm_database("A Book", b"BOOK", b"MOBI");
    header.extend_from_slice(b"MOBI");
    header
}

/// The front of a file whose magic sits an offset into it rather than at its front — a disk
/// image's, or the volume signature of a Macintosh one — as the bytes a probe would hold.
fn at_offset(offset: usize, magic: &[u8]) -> Vec<u8> {
    let mut probe = vec![0u8; offset + magic.len()];
    probe[offset..].copy_from_slice(magic);
    probe
}

/// A probe of `needle` at an offset, for the formats whose marker is not at the front
/// of the file.
fn padded(offset: usize, needle: &[u8]) -> Vec<u8> {
    let mut probe = vec![0u8; offset + needle.len()];
    probe[offset..].copy_from_slice(needle);

    probe
}

/// The head of a package that declares its own type: the first entry of the archive is
/// a stored `mimetype` entry, and the type itself follows the thirty-byte header of
/// that entry and the eight characters of its name.
fn declared_package(mime: &str) -> Vec<u8> {
    let mut probe = vec![0u8; 38];

    probe[..4].copy_from_slice(b"PK\x03\x04");
    probe[26..28].copy_from_slice(&8u16.to_le_bytes());
    probe[30..38].copy_from_slice(b"mimetype");
    probe.extend_from_slice(mime.as_bytes());
    // The next record of the archive, which is what a reader of the declaration finds
    // where the type ends.
    probe.extend_from_slice(b"PK\x03\x04");

    probe
}

mod answers;
mod table;
