//! Which files are images.
//!
//! The extension list lives in `config.ini`, written from the built-in list on
//! first run and read back from there, exactly as the text preview's lists and the
//! archive and office lists are — so a user can add a format this list does not
//! name, or take one out, without a rebuild.
//!
//! The question here is only what a file is *called*. Nothing is probed: the
//! decoder has the last word on whether a file is an image at all, and it is handed
//! one only after the hover gate has said this is the kind of file it is, so a file
//! refused here is never opened.

use crate::text_formats;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the still and animated
/// picture formats a hover is expected to meet.
///
/// `apng` is a name like any other here — an animated PNG is recognized by its own
/// `acTL` chunk rather than by what it is called — and `gif`, `png` and `webp` are
/// each both an animated format and a still one.
pub const DEFAULT_IMAGE_EXTENSIONS: &str =
    "jpg,jpeg,jpe,jfif,png,apng,gif,bmp,ico,tiff,tif,webp,tga,pbm,pgm,ppm,pam,pnm,hdr,exr,qoi,ff";

/// Read one entry out of the configured list into the lowercase form the lookups
/// use.
///
/// Every picture format is a bare extension — unlike the archive list, which has to
/// carry the dotted `tar.gz` — so anything that is not one is dropped rather than
/// matched against.
pub fn sanitize_image_extensions(list: &str) -> Vec<String> {
    let mut extensions: Vec<String> = Vec::new();

    for entry in list.split(',') {
        let trimmed = entry.trim().trim_start_matches('.').to_lowercase();
        let is_extension = !trimmed.is_empty()
            && !trimmed.starts_with('.')
            && !trimmed.ends_with('.')
            && trimmed
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '-' | '_'));

        if is_extension && !extensions.contains(&trimmed) {
            extensions.push(trimmed);
        }
    }

    extensions
}

/// Whether the configured list claims `path`.
pub fn matches_image_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}
