use std::fs::File;
use std::io::Read;
use std::path::Path;

const MPEGTS_SYNC_BYTE: u8 = 0x47;
const MPEGTS_PACKET_SIZES: [usize; 3] = [188, 192, 204];
const MPEGTS_SYNC_RUN: usize = 4;
const MPEGTS_PROBE_BYTES: usize = 2048;

// TypeScript also uses these two extensions, so they need a content check
const TYPESCRIPT_SHARED_EXTENSIONS: &[&str] = &["ts", "mts"];

/// Every container and raw video stream FFmpeg is able to demux.
pub const VIDEO_EXTENSIONS: &[&str] = &[
    "mp4", "m4v", "mov", "qt", "3gp", "3g2", "3gpp", "mj2", "psp", "ismv", "f4v", "mkv", "mk3d",
    "webm", "ts", "m2t", "m2ts", "mts", "tr", "tp", "tod", "wtv", "dvr-ms", "ty", "ty+", "mpg",
    "mpeg", "mpe", "mpv", "m1v", "m2v", "m2p", "vob", "vro", "h261", "h263", "h264", "h26l", "264",
    "avc", "h265", "hevc", "265", "h266", "vvc", "266", "vc1", "rcv", "av1", "obu", "evc", "apv",
    "avs", "avs2", "avs3", "cavs", "drc", "vc2", "y4m", "ivf", "avi", "divx", "asf", "wmv", "rm",
    "rmvb", "flv", "swf", "ogv", "ogm", "mxf", "gxf", "dv", "dif", "nut", "nsv", "mjpg", "mjpeg",
    "bik", "bk2", "smk", "roq", "mve", "cpk", "thp", "usm", "moflex", "xmv", "mvi", "mxg", "rsd",
    "str", "cin", "c93", "cdxl", "xl", "flm", "yop", "imx", "dav", "viv", "ivr", "vw", "cdg",
    "pmp", "kux", "ifv",
];

pub fn is_video_file(path: &Path) -> bool {
    let extension = match path.extension().and_then(|ext| ext.to_str()) {
        Some(extension) => extension.to_lowercase(),
        None => return false,
    };

    if !VIDEO_EXTENSIONS.contains(&extension.as_str()) {
        return false;
    }

    if TYPESCRIPT_SHARED_EXTENSIONS.contains(&extension.as_str()) {
        return looks_like_mpegts(path);
    }

    true
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

fn has_mpegts_packets(probe: &[u8]) -> bool {
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
    use super::{has_mpegts_packets, is_video_file, VIDEO_EXTENSIONS};
    use std::path::Path;

    fn packets(size: usize, start: usize, count: usize) -> Vec<u8> {
        let mut probe = vec![0u8; start + size * count];
        for packet in 0..count {
            probe[start + packet * size] = 0x47;
        }
        probe
    }

    #[test]
    fn extensions_are_lowercase_and_unique() {
        let mut sorted = VIDEO_EXTENSIONS.to_vec();
        sorted.sort_unstable();
        sorted.dedup();
        assert_eq!(sorted.len(), VIDEO_EXTENSIONS.len());
        assert!(VIDEO_EXTENSIONS
            .iter()
            .all(|extension| *extension == extension.to_lowercase()));
    }

    #[test]
    fn transport_stream_layouts_are_recognized() {
        assert!(has_mpegts_packets(&packets(188, 0, 5)));
        assert!(has_mpegts_packets(&packets(192, 4, 5)));
        assert!(has_mpegts_packets(&packets(204, 0, 5)));
        assert!(has_mpegts_packets(&packets(188, 137, 5)));
        assert!(has_mpegts_packets(&packets(188, 0, 100)));
    }

    #[test]
    fn source_code_is_not_a_transport_stream() {
        let source = "export const preview = { width: 1280, height: 720 };\n".repeat(64);
        assert!(!has_mpegts_packets(source.as_bytes()));
        assert!(!has_mpegts_packets(&[0u8; 4096]));
        assert!(!has_mpegts_packets(b""));
        assert!(!has_mpegts_packets(&packets(188, 0, 3)));
    }

    #[test]
    fn extension_gate_matches_supported_files() {
        assert!(is_video_file(Path::new("clip.mp4")));
        assert!(is_video_file(Path::new("clip.MKV")));
        assert!(is_video_file(Path::new("recording.m2ts")));
        assert!(!is_video_file(Path::new("notes.txt")));
        assert!(!is_video_file(Path::new("clip")));
        assert!(!is_video_file(Path::new("missing.ts")));
    }
}
