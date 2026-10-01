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

/// The extensions written to `config.ini` on first run under `[video]`: the containers and raw
/// streams the codecs Windows 11 itself has can demux and decode — the ISO base media family, AVI,
/// ASF, Matroska and WebM, and the MPEG-1, MPEG-2 and MPEG-4 elementary, program and transport
/// streams. It is the same list the README's *Videos Windows 11 plays itself* row names, and the
/// two are meant to be read side by side: what is here is what a machine with no FFmpeg on it
/// still plays.
///
/// Which list a name is in is which engine a video is played by, and that is what decides what a
/// pinned window of one can do: a file of these names is played by the media engine Windows has,
/// in this app's own window, so its frames are this app's to draw — a pinned one is resized by its
/// edges, maximized by its caption and dragged by its picture, and its transport bar is a real
/// control. A name in `[ffmpeg]` beside it is played by FFmpeg's `ffplay` in a window of its own
/// instead, which this app cannot resize, seek or pause.
///
/// What is *not* claimed here is that the machine in hand decodes every file of one of these
/// names: a `.mkv` of HEVC on a machine with no HEVC codec, or an `.mp4` of ProRes, is a file the
/// engine is asked about and turns down. That question is asked of the engine itself, once per
/// file and version (`video_player::plays`), and a file it turns down is played by FFmpeg's player
/// where one is installed — the list decides which engine is asked first, not which engine ends up
/// playing (see `preview_window::media_engine_plays`).
pub const DEFAULT_VIDEO_EXTENSIONS: &str = "3g2,3gp,3gpp,asf,avi,dvr-ms,m1v,m2t,m2ts,m2v,m4v,mkv,\
mov,mp4,mpe,mpeg,mpg,mts,qt,ts,vob,webm,wmv";

/// The extensions written to `config.ini` on first run under `[ffmpeg]`: every container and raw
/// video stream FFmpeg is able to demux that Windows' own codecs are not asked about — the
/// streams no decoder of Windows' reads (AV1, VC-1, Flash's formats), the containers whose handler
/// Windows does not ship (RealMedia, MXF, NUT, the game and camera formats), and the names the
/// ISO base media family is shared with where what is inside is not what Windows decodes.
///
/// It is the README's *Needs FFmpeg* list, and it is the rest of the one list the two were: a name
/// here is played by FFmpeg's `ffplay`, which is the engine that plays everything. A machine with
/// no FFmpeg installed shows nothing for one of these names rather than being asked about it —
/// that is what the list is for. Moving a name from here to `[video]` is the whole of asking the
/// media engine about it instead, and a name the engine cannot open costs one probe and falls back
/// to the player anyway (see `DEFAULT_VIDEO_EXTENSIONS`).
pub const DEFAULT_FFMPEG_EXTENSIONS: &str =
    "264,265,266,apv,av1,avc,avs,avs2,avs3,bik,bk2,c93,cavs,cdg,cdxl,cin,cpk,dav,\
dif,divx,drc,dv,evc,f4v,flm,flv,gxf,h261,h263,h264,h265,h266,h26l,hevc,ifv,imx,ismv,ivf,ivr,\
kux,m2p,mj2,mjpeg,mjpg,mk3d,moflex,mpv,mve,mvi,mxf,mxg,nsv,nut,obu,ogm,ogv,pmp,psp,rcv,rm,rmvb,\
roq,rsd,smk,str,swf,thp,tod,tp,tr,ty,ty+,usm,vc1,vc2,viv,vro,vvc,vw,wtv,xl,xmv,y4m,yop";

/// The one list the two above were one list of: every container and raw video stream FFmpeg is
/// able to demux, which is what every build before the split wrote under `[video]`.
///
/// It is here for the same reason the `*_BEFORE_*` lists beside the other kinds are: a list is
/// only ever read out of `config.ini` — nothing in the tray edits one — so a `[video]` list
/// holding exactly these entries is this app's own older list rather than an edit somebody made
/// by hand, and it is what tells `repair_older_lists` that a file written before the split is to
/// be split rather than kept whole. A list with any entry added, removed or spelled differently
/// is the user's and is left exactly as it is.
pub const VIDEO_EXTENSIONS_BEFORE_THE_SPLIT: &str = "264,265,266,3g2,3gp,3gpp,apv,asf,av1,avc,avi,avs,avs2,avs3,bik,bk2,c93,cavs,cdg,cdxl,cin,cpk,dav,\
dif,divx,drc,dv,dvr-ms,evc,f4v,flm,flv,gxf,h261,h263,h264,h265,h266,h26l,hevc,ifv,imx,ismv,ivf,\
ivr,kux,m1v,m2p,m2t,m2ts,m2v,m4v,mj2,mjpeg,mjpg,mk3d,mkv,moflex,mov,mp4,mpe,mpeg,mpg,mpv,mts,mve,\
mvi,mxf,mxg,nsv,nut,obu,ogm,ogv,pmp,psp,qt,rcv,rm,rmvb,roq,rsd,smk,str,swf,thp,tod,tp,tr,ts,ty,\
ty+,usm,vc1,vc2,viv,vob,vro,vvc,vw,webm,wmv,wtv,xl,xmv,y4m,yop";

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
    if !extensions.iter().any(|claimed| claimed.as_str() == extension) {
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
    extensions.iter().any(|claimed| claimed.as_str() == extension)
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
        let before = crate::formats::text_formats::sanitize_extension_list(VIDEO_EXTENSIONS_BEFORE_THE_SPLIT);
        let native = crate::formats::text_formats::sanitize_extension_list(DEFAULT_VIDEO_EXTENSIONS);
        let ffmpeg = crate::formats::text_formats::sanitize_extension_list(DEFAULT_FFMPEG_EXTENSIONS);

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
            DEFAULT_VIDEO_EXTENSIONS,
            DEFAULT_FFMPEG_EXTENSIONS,
            VIDEO_EXTENSIONS_BEFORE_THE_SPLIT,
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
