//! Files whose content is not on this machine yet.
//!
//! OneDrive, SharePoint, Dropbox and the other synchronizing providers leave a
//! placeholder in the folder — a few kilobytes of metadata carrying the file's
//! name, size and dates — and pull the content down the first time something
//! reads it. Nothing about the file says so: an ordinary open is what starts the
//! transfer, and the reader simply waits for it.
//!
//! That wait is not one a hover can pay. It is unbounded, it is a real download
//! the user did not ask for, and on a metered connection it is a bill. A
//! placeholder is therefore answered with no preview rather than with a read.

use std::os::windows::fs::MetadataExt;
use std::path::Path;

/// The attributes that mean the content lives somewhere else and reading the
/// file is what fetches it.
///
/// The two recall flags are what the cloud providers set: `RECALL_ON_OPEN` on a
/// file whose handle has to be reparse-pointed, `RECALL_ON_DATA_ACCESS` on the
/// ordinary placeholder whose data is paged in on first access. `OFFLINE` covers
/// files on tiered and remotely-stored volumes, which the same read would recall.
const FILE_ATTRIBUTE_OFFLINE: u32 = 0x0000_1000;
const FILE_ATTRIBUTE_RECALL_ON_OPEN: u32 = 0x0004_0000;
const FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS: u32 = 0x0040_0000;

const REMOTE_ATTRIBUTES: u32 =
    FILE_ATTRIBUTE_OFFLINE | FILE_ATTRIBUTE_RECALL_ON_OPEN | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS;

/// Whether reading `path` would have to fetch its content first.
///
/// The attributes are read rather than the file, because asking for a file's
/// attributes is a question about its directory entry: it is the same question
/// the shell asks to draw the cloud state in a file's icon, and unlike opening
/// the file it does not start the download.
pub fn needs_download(path: &Path) -> bool {
    let Ok(metadata) = std::fs::metadata(path) else {
        return false;
    };

    is_remote(metadata.file_attributes())
}

/// The same question asked of attributes a caller has already read: one directory entry
/// answers whether a file is there, what version it is at and whether its content is on this
/// machine, and a hover that has read one is not made to read another (see
/// `crate::formats::head::Facts`).
pub(crate) fn is_remote(attributes: u32) -> bool {
    attributes & REMOTE_ATTRIBUTES != 0
}
