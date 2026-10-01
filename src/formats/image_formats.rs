// Which files are images.
//
// The list of names this answers for is a row of `crate::formats::lists` — the one table every
// kind's list is a row of, and the one place a list is written down. What is left here is the
// question the list cannot answer, which is the same question every other kind's file asks.
//
// Nothing is probed: the question here is only what a file is *called*. The decoder has the
// last word on whether a file is an image at all, and it is handed one only after the hover
// gate has said this is the kind of file it is, so a file refused here is never opened.

use std::path::Path;

/// Whether the configured list claims `path`.
pub fn matches_image_list(path: &Path, extensions: &[String]) -> bool {
    crate::formats::text_formats::matches_configured_extension(path, extensions)
}
