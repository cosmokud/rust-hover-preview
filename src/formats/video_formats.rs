//! Which files are videos, and the one question a name cannot settle on its own.
//!
//! A video's names are two lists rather than one — `[video]` for what the media engine Windows
//! has is asked to play, `[ffmpeg]` for what FFmpeg's player is — and both are rows of
//! `crate::formats::lists` beside every other kind's. What is kept here is the part of the
//! question the two lists cannot answer between them: `.ts` and `.mts` are claimed by the text
//! lists too, so a file of either name is settled by its own bytes, and the two lists are asked
//! together wherever the question is what a file *is* rather than which engine plays it.

use crate::config::config::{AppConfig, PreviewType};
use crate::CONFIG;
use std::fs::File;
use std::io::Read;
use std::path::Path;

const MPEGTS_SYNC_BYTE: u8 = 0x47;
const MPEGTS_PACKET_SIZES: [usize; 3] = [188, 192, 204];
const MPEGTS_SYNC_RUN: usize = 4;
const MPEGTS_PROBE_BYTES: usize = 2048;

// TypeScript also uses these two extensions, so they need a content check
const TYPESCRIPT_SHARED_EXTENSIONS: &[&str] = &["ts", "mts"];

/// Whether one of the two video lists claims `path` as a video.
///
/// `.ts` and `.mts` are claimed by the text lists too, so a file of either name is
/// settled by its content: an MPEG-TS sync byte makes it the transport stream the
/// list says it is, and a file without one falls through to the text preview of the
/// TypeScript source it is.
///
/// The two lists are asked together wherever the question is what a file *is* — which is
/// every question but the engine's own: `[video]` and `[ffmpeg]` are what a video is played
/// by, not what it is, so a name in either is a video to the router, the layout and the
/// engines that exclude the video names from their own lists.
pub fn matches_video_list(path: &Path, extensions: &[String]) -> bool {
    let Some(extension) = crate::formats::text_formats::lookup_extension(path) else {
        return false;
    };

    list_matches_video(&extension, path, extensions)
}

/// Whether one list claims a name already looked up, asking the file only for the two names the
/// text lists share with these (see [`matches_video_list`]).
fn list_matches_video(extension: &str, path: &Path, extensions: &[String]) -> bool {
    if !extensions
        .iter()
        .any(|claimed| claimed.as_str() == extension)
    {
        return false;
    }

    if TYPESCRIPT_SHARED_EXTENSIONS.contains(&extension) {
        return looks_like_mpegts(path);
    }

    true
}

/// Whether one list claims a name already looked up, asking nothing of the file: the half of
/// [`list_matches_video`] that does not settle the two shared extensions by their content.
fn list_claims_video_name(extension: &str, extensions: &[String]) -> bool {
    extensions
        .iter()
        .any(|claimed| claimed.as_str() == extension)
}

/// Whether either of the two lists claims `path`, asked of a configuration in hand: what every
/// caller that already holds the configuration asks rather than the one above, which is the
/// question with the global read out of it.
pub fn matches_any_video_list(path: &Path, config: &AppConfig) -> bool {
    // The name is read once for both lists: a lookup is an allocation, and this is the question
    // every caller without an answer of its own asks (see `lookup_extension`).
    let Some(extension) = crate::formats::text_formats::lookup_extension(path) else {
        return false;
    };

    list_matches_video(&extension, path, &config.video_extensions)
        || list_matches_video(&extension, path, &config.ffmpeg_extensions)
}

/// As above, by name alone: the half of [`matches_any_video_list`] that asks nothing of the file.
pub fn claims_any_video_name(path: &Path, config: &AppConfig) -> bool {
    // The name is read once for both lists, as above (see `lookup_extension`).
    let Some(extension) = crate::formats::text_formats::lookup_extension(path) else {
        return false;
    };

    list_claims_video_name(&extension, &config.video_extensions)
        || list_claims_video_name(&extension, &config.ffmpeg_extensions)
}

/// Whether the configured lists claim `path`, without asking whether video previews
/// are switched on.
pub fn is_video_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_any_video_list(path, &config))
        .unwrap_or(false)
}

/// Whether a video preview may be shown for `path`: the file the configured list
/// claims, and the `Videos` gate in the tray's `Preview Types` submenu.
///
/// [`is_video_file`] is the classification on its own, which is what asks whether
/// a file is a video rather than whether one may be shown — a `.ts` a video gate
/// turned down must still be recognized as the transport stream it is, or the
/// text lists would claim it.
pub fn is_video_preview(path: &Path) -> bool {
    is_video_file(path) && PreviewType::Videos.enabled()
}

fn looks_like_mpegts(path: &Path) -> bool {
    let mut file = match File::open(path) {
        Ok(file) => file,
        Err(_) => return false,
    };

    let mut probe = [0u8; MPEGTS_PROBE_BYTES];
    match file.read(&mut probe) {
        Ok(read) => has_mpegts_packets(&probe[..read]),
        Err(_) => false,
    }
}

pub(crate) fn has_mpegts_packets(probe: &[u8]) -> bool {
    MPEGTS_PACKET_SIZES.iter().any(|&size| {
        let span = size * (MPEGTS_SYNC_RUN - 1);
        if probe.len() <= span {
            return false;
        }

        let last_start = (probe.len() - span - 1).min(size - 1);
        (0..=last_start).any(|start| {
            (0..MPEGTS_SYNC_RUN).all(|packet| probe[start + packet * size] == MPEGTS_SYNC_BYTE)
        })
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The split is a partition: every name the one list held is in exactly one of the two, so a
    /// `config.ini` written before it loses no format and asks the media engine about nothing it
    /// was not already asked about.
    #[test]
    fn the_two_video_lists_partition_the_one_list_they_were() {
        let before = crate::formats::text_formats::sanitize_extension_list(
            crate::formats::lists::VIDEO_EXTENSIONS_BEFORE_THE_SPLIT,
        );
        let native = crate::formats::text_formats::sanitize_extension_list(
            crate::formats::lists::DEFAULT_VIDEO_EXTENSIONS,
        );
        let ffmpeg = crate::formats::text_formats::sanitize_extension_list(
            crate::formats::lists::DEFAULT_FFMPEG_EXTENSIONS,
        );

        for name in &native {
            assert!(!ffmpeg.contains(name), "{name} is in both video lists");
            assert!(before.contains(name), "{name} is in neither video list");
        }

        for name in &before {
            assert!(
                native.contains(name) || ffmpeg.contains(name),
                "{name} was dropped by the split"
            );
        }

        assert_eq!(
            native.len() + ffmpeg.len(),
            before.len(),
            "every name belongs to exactly one of the two lists"
        );
    }

    /// Both lists are written in the order a file is read in, which is what lets the repair that
    /// splits a list this app wrote compare the two entry for entry.
    #[test]
    fn the_video_lists_are_written_in_order() {
        for list in [
            crate::formats::lists::DEFAULT_VIDEO_EXTENSIONS,
            crate::formats::lists::DEFAULT_FFMPEG_EXTENSIONS,
            crate::formats::lists::VIDEO_EXTENSIONS_BEFORE_THE_SPLIT,
        ] {
            let entries = crate::formats::text_formats::sanitize_extension_list(list);
            let mut sorted = entries.clone();
            sorted.sort();

            assert_eq!(entries, sorted, "the list is not in alphabetical order");
        }
    }

    /// Which of the two a name belongs to is which engine plays it, so the names that could be
    /// argued either way are pinned here: the containers and streams Windows' own codecs read,
    /// and the ones only FFmpeg's player does.
    #[test]
    fn the_two_lists_put_each_engine_where_it_belongs() {
        let config = AppConfig::default();
        let native = |name: &str| matches_video_list(Path::new(name), &config.video_extensions);

        // `ts` and `mts` are left out of this: they are claimed by the text lists too, so a file
        // of either name is settled by its own bytes rather than by the list it is in (see
        // `looks_like_mpegts`), and there is no file here to read.
        for name in [
            "film.mp4",
            "clip.mkv",
            "movie.webm",
            "old.avi",
            "clip.mpg",
            "tape.wmv",
            "disc.vob",
        ] {
            assert!(native(name), "{name} is played by the media engine");
        }

        for name in [
            "clip.flv",
            "show.ogv",
            "song.wtv",
            "raw.y4m",
            "capture.dv",
            "film.hevc",
        ] {
            assert!(!native(name), "{name} is played by FFmpeg's player");
        }
    }

    /// Either list makes a file a video: the two are which engine plays one, and nothing else
    /// asks which of them carries a name.
    #[test]
    fn a_name_in_either_list_is_a_video() {
        let config = AppConfig::default();

        for name in ["film.mkv", "clip.mp4", "old.flv", "raw.y4m"] {
            assert!(
                matches_any_video_list(Path::new(name), &config),
                "{name} is a video"
            );
            assert!(
                claims_any_video_name(Path::new(name), &config),
                "{name} is a video by name"
            );
        }

        assert!(!matches_any_video_list(Path::new("notes.txt"), &config));
        assert!(!claims_any_video_name(Path::new("notes.txt"), &config));
    }
}
