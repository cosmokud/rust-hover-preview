//! Project containers, read for the picture they hold of the whole document.
//!
//! A Krita document, an OpenRaster one and a Procreate one are zip containers that keep a
//! picture of the finished work beside the layers it is built from: Krita writes the
//! flattened document as `mergedimage.png` and a small preview as `preview.png`,
//! OpenRaster writes the same flattened document as `mergedimage.png` and a file-manager
//! thumbnail under `Thumbnails/`, and a Procreate document keeps the thumbnail a file
//! manager shows on iPad under `QuickLook/`. So the preview here is not drawn and not
//! composited: it is the picture the application itself saved, read out of the container
//! and shown.
//!
//! What that picture is worth differs from container to container and is worth being
//! plain about. Some of them hold the *document*: a merged image is the flattened
//! picture at the size the work was made at, so a preview of one is the document itself
//! at the size the display has room for. The others hold a thumbnail: a few hundred
//! pixels the application wrote for a file manager, which is a preview that tells a user
//! which file this is and no more — and one of those is a small picture laid out at the
//! share of the display every other design preview is laid out at, so a thumbnail at
//! `Fit to Screen` is a thumbnail enlarged. That is why the document is preferred by name
//! wherever a container holds both, and why a container that holds only a thumbnail is
//! still read: a small picture of the right drawing beats no picture at all.
//!
//! A container none of those names answers for is a file this app has no reader for,
//! and its answer is no preview, like any other format that is not read here. Nothing
//! is unpacked to disk: a member is inflated into memory under the decode budget and
//! decoded from there.

use crate::config::config::{decode_budget_bytes, image_decode_limits};
use image::GenericImageView;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;
use zip::ZipArchive;

/// The names a container keeps the picture of the whole document under, in the order
/// they are worth reading: the flattened document first, since it is the size the
/// work was made at, and the previews an application writes for a file manager after
/// it, since one of those is a thumbnail rather than the document.
///
/// The names are the ones the formats themselves fix — Krita's and OpenRaster's
/// `mergedimage.png`, OpenRaster's `Thumbnails/thumbnail.png`, a Procreate document's
/// `QuickLook/Thumbnail.png`, and the two bitmaps a CorelDRAW document writes beside
/// itself under `metadata/thumbnails/` — and the rest are the same picture written by
/// applications that keep a container of their own. Whichever of them a file holds is
/// the one read, and the comparison is by name whatever case it is written in, since
/// a container is written by an application rather than by this app.
///
/// A picture of the whole document comes before the names a container may keep a picture
/// under without saying what it is: the flattened document a project keeps is the document
/// itself, and what is left is the picture the application wrote for a file manager.
const PREVIEW_MEMBERS: &[&str] = &[
    // The flattened document, at the size the work was made at.
    "mergedimage.png",
    // The picture the application wrote for a file manager.
    "Thumbnails/thumbnail.png",
    "QuickLook/Thumbnail.png",
    // And the names a container may keep a preview under without saying which it is.
    "previews/preview.png",
    "thumbnail.png",
    "preview.png",
];

/// How much of a member is read to find out the size of the picture it holds. A PNG
/// says its size in its first few bytes and a JPEG a little way in, so what is read
/// is a header's worth of the member rather than the member.
const DIMENSION_PROBE_BYTES: u64 = 1 << 20;

/// Entries the size cache holds before it is emptied.
const DIMENSION_CACHE_MAX_ENTRIES: usize = 512;

/// What a size is only valid for: the container, and the version of it the size was
/// read from. A project saved again is another document — a file saved with another
/// layer shown is the case that matters — and its preview is another picture.
#[derive(Clone, PartialEq, Eq, Hash, Debug)]
struct DimensionKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

/// A picture's size, keyed by the container and its version. `None` records a file
/// with no preview member in it, so a container this reader has no answer for is not
/// opened again on every hover.
type DimensionCache = HashMap<DimensionKey, Option<(u32, u32)>>;

static DIMENSIONS: Lazy<Mutex<DimensionCache>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// The size of the picture `path` holds a preview of.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let key = dimension_key(path);
    if let Ok(cache) = DIMENSIONS.lock() {
        if let Some(cached) = cache.get(&key) {
            return *cached;
        }
    }

    let dimensions = probe_dimensions(path);
    remember_dimensions(key, dimensions);

    dimensions
}

/// The picture `path` holds, decoded into `width` by `height` and handed back as BGRA,
/// top down, the way every frame in this app is composed.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    let mut archive = ZipArchive::new(File::open(path).ok()?).ok()?;
    let index = preview_member(&mut archive)?;
    let bytes = member_bytes(&mut archive, index)?;

    let mut reader =
        image::ImageReader::new(std::io::Cursor::new(&bytes)).with_guessed_format().ok()?;
    reader.limits(image_decode_limits());
    let picture = reader.decode().ok()?;

    let (source_width, source_height) = picture.dimensions();
    let picture = if source_width != width || source_height != height {
        picture.resize_exact(width, height, image::imageops::FilterType::Triangle)
    } else {
        picture
    };

    Some(crate::ui::preview_window::rgba_to_bgra(picture.to_rgba8().as_raw()))
}

/// The size the picture is stored at, read from the picture's own header.
fn probe_dimensions(path: &Path) -> Option<(u32, u32)> {
    let mut archive = ZipArchive::new(File::open(path).ok()?).ok()?;
    let index = preview_member(&mut archive)?;
    let entry = archive.by_index(index).ok()?;

    let take = entry.size().min(DIMENSION_PROBE_BYTES);
    let mut header = Vec::new();
    entry.take(take).read_to_end(&mut header).ok()?;

    let reader = image::ImageReader::new(std::io::Cursor::new(&header))
        .with_guessed_format()
        .ok()?;

    reader.into_dimensions().ok()
}

/// The member a container's preview is read from: the name of the greatest worth that
/// this container holds.
///
/// The names are asked in their own order rather than the container's, because the
/// order a container writes its members in is its own business: an OpenRaster file
/// that holds its thumbnail before its merged image — which is what writing the names
/// alphabetically comes to — is a document whose preview would otherwise be the
/// thumbnail rather than the document.
fn preview_member(archive: &mut ZipArchive<File>) -> Option<usize> {
    let count = archive.len();

    for candidate in PREVIEW_MEMBERS {
        for index in 0..count {
            let found = archive
                .by_index_raw(index)
                .map(|entry| entry.name().eq_ignore_ascii_case(candidate))
                .unwrap_or(false);

            if found {
                return Some(index);
            }
        }
    }

    None
}

/// A member's bytes, read under the decode budget the rest of this app reads files
/// under: what a member declares is what it may ask for, and a member that declares
/// more than a hover may hold is one that is not read at all.
fn member_bytes(archive: &mut ZipArchive<File>, index: usize) -> Option<Vec<u8>> {
    let entry = archive.by_index(index).ok()?;
    let budget = decode_budget_bytes();

    if entry.size() > budget {
        return None;
    }

    let mut bytes = Vec::new();
    entry.take(budget + 1).read_to_end(&mut bytes).ok()?;

    (bytes.len() as u64 <= budget).then_some(bytes)
}

/// The container and the version of it a size is read from, read the way every other
/// held value in this app is: what a file is, is its name as it is now, what it
/// weighed, and when it was last written.
fn dimension_key(path: &Path) -> DimensionKey {
    let metadata = std::fs::metadata(path).ok();

    DimensionKey {
        path: path.to_path_buf(),
        modified: metadata
            .as_ref()
            .and_then(|metadata| metadata.modified().ok()),
        len: metadata.map(|metadata| metadata.len()).unwrap_or(0),
    }
}

/// Record a size that has been read, so a container's table of contents is not walked
/// again for every hover.
fn remember_dimensions(key: DimensionKey, dimensions: Option<(u32, u32)>) {
    if let Ok(mut cache) = DIMENSIONS.lock() {
        if !cache.contains_key(&key) && cache.len() >= DIMENSION_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, dimensions);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use image::ImageEncoder;
    use std::io::Write;

    /// A container holding the members named, in the order they are named, stored rather
    /// than deflated: what these fixtures are asked is which member is chosen.
    fn write_container(path: &Path, members: &[(&str, Vec<u8>)]) {
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let mut writer = zip::ZipWriter::new(File::create(path).expect("a zip to write"));

        for (name, bytes) in members {
            writer.start_file(*name, options).expect("a member");
            writer.write_all(bytes).expect("bytes");
        }

        writer.finish().expect("a finished zip");
    }

    fn bmp(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = Vec::new();
        image::codecs::bmp::BmpEncoder::new(&mut bytes)
            .write_image(
                &vec![0u8; (width * height * 3) as usize],
                width,
                height,
                image::ExtendedColorType::Rgb8,
            )
            .expect("a bitmap");

        bytes
    }

    /// The names are asked in the order they are worth reading rather than in the order the
    /// container happens to write them in, and that order puts a flattened document ahead of
    /// every thumbnail a container may hold beside it.
    #[test]
    fn reads_the_document_rather_than_the_thumbnail() {
        let path = std::env::temp_dir().join("rust-hover-preview-container-order");

        for (picture, other) in [
            ("mergedimage.png", "Thumbnails/thumbnail.png"),
            ("mergedimage.png", "QuickLook/Thumbnail.png"),
            ("mergedimage.png", "previews/preview.png"),
        ] {
            write_container(&path, &[(other, bmp(3, 5)), (picture, bmp(6, 7))]);

            assert_eq!(
                dimensions(&path),
                Some((6, 7)),
                "`{picture}` is the picture read, though `{other}` is written first"
            );
        }

        std::fs::remove_file(&path).ok();
    }

    /// A container holding nothing but the thumbnail its application wrote for a file
    /// manager is read for that: a small picture of the right drawing is a preview where
    /// none at all would otherwise be.
    #[test]
    fn reads_the_thumbnail_when_it_is_all_the_container_holds() {
        let path = std::env::temp_dir().join("rust-hover-preview-container-thumbnail");
        write_container(&path, &[("QuickLook/Thumbnail.png", bmp(4, 6))]);

        assert_eq!(dimensions(&path), Some((4, 6)));

        std::fs::remove_file(&path).ok();
    }
}
