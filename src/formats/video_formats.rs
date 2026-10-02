//! Which files are videos, and the one question a name cannot settle on its own.
//!
//! A video's names are two lists rather than one — `[video]` for what the media engine Windows
//! has is asked to play, `[ffmpeg]` for what FFmpeg's player is — and both are rows of
//! `crate::formats::lists` beside every other kind's. What is kept here is the part of the
//! question the two lists cannot answer between them: `.ts` and `.mts` are claimed by the text
//! lists too, so a file of either name is settled by its own bytes, and the two lists are asked
//! together wherever the question is what a file *is* rather than which engine plays it.

use crate::config::config::AppConfig;
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

/// Whether one list claims a name already looked up, asking the file only for the two names the
/// text lists share with these (see [`matches_any_video_list`]).
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

/// Whether either of the two lists claims `path` as a video, asked of a configuration in hand.
///
/// `.ts` and `.mts` are claimed by the text lists too, so a file of either name is settled by its
/// content: an MPEG-TS sync byte makes it the transport stream the list says it is, and a file
/// without one falls through to the text preview of the TypeScript source it is.
///
/// The two lists are asked together wherever the question is what a file *is* — which is every
/// question but the engine's own: `[video]` and `[ffmpeg]` are what a video is played by, not what
/// it is, so a name in either is a video to the router, the layout and the engines that exclude
/// the video names from their own lists.
///
/// It reads the two rows of the table rather than two fields of the configuration, because the
/// rows are what own them, and because this is the one function in the app that opens a file from
/// inside a kind's list question: a caller holding the configuration's lock must not ask it (see
/// `routing::named_as`, which is what `preview_window` asks instead).
pub fn matches_any_video_list(path: &Path, config: &AppConfig) -> bool {
    // The name is read once for both lists: a lookup is an allocation, and this is the question
    // every caller without an answer of its own asks (see `lookup_extension`).
    let Some(extension) = crate::formats::text_formats::lookup_extension(path) else {
        return false;
    };

    list_matches_video(
        &extension,
        path,
        crate::formats::lists::VIDEO.entries(config),
    ) || list_matches_video(
        &extension,
        path,
        crate::formats::lists::FFMPEG.entries(config),
    )
}

/// [`matches_any_video_list`] of two lists the caller already holds.
///
/// This is the one video question that opens a file — `.ts` and `.mts` are settled by their own
/// bytes — so it is also the one form that cannot be asked with the configuration's lock in hand,
/// and the caller that has no list of its own has to copy these two out from under the guard and
/// let it go before it gets here. A guard held across that read is a guard the hook's 15 ms tick,
/// the tray, the preview thread's own message pump and every engine wait on for as long as the
/// volume takes to answer it (see `preview_window::HoverFacts::read`, which copies the same two
/// out for the same reason, and `routing::named_as`, which is the question asked when the
/// configuration is in hand).
pub fn matches_any_video_list_in(path: &Path, video: &[String], ffmpeg: &[String]) -> bool {
    // The name is read once for both lists, as above (see `lookup_extension`).
    let Some(extension) = crate::formats::text_formats::lookup_extension(path) else {
        return false;
    };

    list_matches_video(&extension, path, video) || list_matches_video(&extension, path, ffmpeg)
}

/// As above, by name alone: the half of [`matches_any_video_list`] that asks nothing of the file.
pub fn claims_any_video_name(path: &Path, config: &AppConfig) -> bool {
    claims_any_video_name_in(
        path,
        crate::formats::lists::VIDEO.entries(config),
        crate::formats::lists::FFMPEG.entries(config),
    )
}

/// [`claims_any_video_name`] of two lists the caller already holds.
///
/// It is the form a caller that has to give the guard up before the question can be asked uses,
/// which is every caller on this side of a `File::open`: this answer asks nothing of the disk, so
/// a caller holding the configuration's lock can reach for [`matches_any_video_list`] and a caller
/// that has copied the lists out from under it can reach for this.
pub fn claims_any_video_name_in(path: &Path, video: &[String], ffmpeg: &[String]) -> bool {
    // The name is read once for both lists, as above (see `lookup_extension`).
    let Some(extension) = crate::formats::text_formats::lookup_extension(path) else {
        return false;
    };

    list_claims_video_name(&extension, video) || list_claims_video_name(&extension, ffmpeg)
}

/// Whether the `[video]` row's own list claims `path`, as a question about a *name* — the two
/// names the text lists share are settled by the name alone, with no read of the file.
///
/// It exists because that is the one shape the video kind has that no row can answer: the caller
/// wants the list's own answer for a file it has not opened and may not open — a synchronizing
/// provider's placeholder is a directory entry that can be answered for but a file that must not
/// be opened, because opening it is what starts the download — and so the two shared names are
/// left to the text lists, which is the same answer [`matches_any_video_list`] gives for a name
/// the content has already spoken for.
pub fn claims_video_name(path: &Path, extensions: &[String]) -> bool {
    let Some(extension) = crate::formats::text_formats::lookup_extension(path) else {
        return false;
    };

    list_claims_video_name(&extension, extensions)
}

/// Whether the configured lists claim `path`, without asking whether video previews
/// are switched on.
///
/// It asks the lists rather than the configuration's rows because it has nothing else to ask them
/// with: the three calls are all on the hook's own 15 ms tick, which reads the configuration into
/// a snapshot of nine scalars and keeps no `AppConfig` to hand down (see `run_explorer_hook`), so
/// a signature that took one would have the tick lock it again on the same line the question is
/// asked on and be nothing but this function's body written somewhere else.
///
/// The guard is let go before the lists are asked, which is the whole of this function. It is the
/// one video question that opens a file — `.ts` and `.mts` are settled by their own bytes — and
/// holding the process-wide lock across that open froze the tray, the preview thread's message pump
/// and every engine behind a `File::open` on one slow volume, asked from the thread that watches
/// Explorer. The two lists are copied out to have somewhere to ask the question from; the copy is
/// two `Vec<String>` on the answer's own path rather than in the tick, which runs sixty times a
/// second (see `preview_window::HoverFacts::read` for the same shape, and
/// [`matches_any_video_list_in`] for the question itself).
///
/// A configuration that will not open answers no, which is the answer the tray being shut gives.
pub fn is_video_file(path: &Path) -> bool {
    let Ok(config) = CONFIG.lock() else {
        return false;
    };

    let video = config.video_extensions.clone();
    let ffmpeg = config.ffmpeg_extensions.clone();
    drop(config);

    matches_any_video_list_in(path, &video, &ffmpeg)
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
    use std::path::PathBuf;

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
        let native = |name: &str| {
            claims_video_name(
                Path::new(name),
                crate::formats::lists::VIDEO.entries(&config),
            )
        };

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

    /// Both forms of the video question answer the same for every name this app ships, and the
    /// two forms are the same question reached two ways.
    ///
    /// The second form takes the two lists rather than the configuration, because it is the one
    /// the file is opened for and a caller holding the configuration's lock cannot be the one
    /// holding it across that open (see [`matches_any_video_list_in`]). What is pinned here is
    /// that handing the lists over changes nothing: every shipped name is answered the same by
    /// both, and a name no list carries is answered no by both.
    ///
    /// What this does *not* say is anything about the lock itself: the discipline is the order of
    /// two statements in [`is_video_file`], which no question asked from outside can see, and the
    /// timing of a slow volume is not a thing a test on this machine can arrange.
    #[test]
    fn both_forms_of_the_question_answer_the_same_for_every_shipped_name() {
        let config = AppConfig::default();
        let video = crate::formats::lists::VIDEO.entries(&config).to_vec();
        let ffmpeg = crate::formats::lists::FFMPEG.entries(&config).to_vec();

        let mut names = video.clone();
        names.extend(ffmpeg.iter().cloned());
        assert!(!names.is_empty(), "this app ships video names");

        // The two shared names are in the list and are left out of the second half of this test:
        // their answer is the file's and not the name's, and there is no file behind a name built
        // out of the list. They have one of their own, with the four files on disk (see
        // `both_forms_settle_the_two_shared_names_by_the_file_behind_them`).
        for name in &names {
            let path = PathBuf::from(format!("clip.{name}"));
            assert_eq!(
                matches_any_video_list(&path, &config),
                matches_any_video_list_in(&path, &video, &ffmpeg),
                "{name} is answered the same by both forms"
            );

            if !TYPESCRIPT_SHARED_EXTENSIONS.contains(&name.as_str()) {
                assert!(
                    matches_any_video_list(&path, &config),
                    "{name} is a name this app ships, and it is a video"
                );
            }
        }

        for name in ["notes.txt", "report.pdf", "archive.zip"] {
            let path = Path::new(name);
            assert!(!matches_any_video_list(path, &config), "{name}");
            assert!(!matches_any_video_list_in(path, &video, &ffmpeg), "{name}");
        }
    }

    /// The two names the video lists share with the text lists are settled by the file's own bytes,
    /// and both forms of the question settle them the same way: a transport stream is a video and
    /// a TypeScript source is not, whichever list it arrives with and whichever form asks.
    ///
    /// These are the two names where handing the lists over could have changed the answer without
    /// any name-only test noticing, because here the answer is not the name's: the file has to be
    /// there to be read, and the read is what [`matches_any_video_list_in`] exists to be able to
    /// do without a guard in hand.
    ///
    /// A real file is behind each of the four, because a synthetic path answers the same either
    /// way — a `File::open` that fails reads as a file without packets, which is the TypeScript
    /// answer and would make the transport stream half of this pass by accident.
    #[test]
    fn both_forms_settle_the_two_shared_names_by_the_file_behind_them() {
        let config = AppConfig::default();
        let video = crate::formats::lists::VIDEO.entries(&config).to_vec();
        let ffmpeg = crate::formats::lists::FFMPEG.entries(&config).to_vec();
        let folder = std::env::temp_dir().join("rust-hover-preview-shared-video-names");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // Four sync bytes one 188-byte packet apart, which is what `has_mpegts_packets` looks
        // for: enough of them to be a transport stream rather than a file that opens with 0x47.
        let mut stream = vec![0u8; MPEGTS_PACKET_SIZES[0] * MPEGTS_SYNC_RUN];
        for packet in 0..MPEGTS_SYNC_RUN {
            stream[packet * MPEGTS_PACKET_SIZES[0]] = MPEGTS_SYNC_BYTE;
        }
        let source: &[u8] = b"export const answer = 42;\n";

        for extension in ["ts", "mts"] {
            for (name, bytes) in [("stream", stream.as_slice()), ("source", source)] {
                let path = folder.join(format!("{name}.{extension}"));
                std::fs::write(&path, bytes).expect("a test file");

                assert_eq!(
                    matches_any_video_list(&path, &config),
                    matches_any_video_list_in(&path, &video, &ffmpeg),
                    "{name}.{extension} is answered the same by both forms"
                );
                assert_eq!(
                    matches_any_video_list(&path, &config),
                    name == "stream",
                    "{name}.{extension} is {}",
                    if name == "stream" {
                        "the transport stream"
                    } else {
                        "the TypeScript source it shares its name with"
                    }
                );
            }
        }

        let _ = std::fs::remove_dir_all(&folder);
    }
}
