use crate::config::PreviewType;
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

/// The extensions written to `config.ini` on first run: every container and raw
/// video stream FFmpeg is able to demux.
pub const DEFAULT_VIDEO_EXTENSIONS: &str = "264,265,266,3g2,3gp,3gpp,apv,asf,av1,avc,avi,avs,avs2,avs3,bik,bk2,c93,cavs,cdg,cdxl,cin,cpk,dav,\
dif,divx,drc,dv,dvr-ms,evc,f4v,flm,flv,gxf,h261,h263,h264,h265,h266,h26l,hevc,ifv,imx,ismv,ivf,\
ivr,kux,m1v,m2p,m2t,m2ts,m2v,m4v,mj2,mjpeg,mjpg,mk3d,mkv,moflex,mov,mp4,mpe,mpeg,mpg,mpv,mts,mve,\
mvi,mxf,mxg,nsv,nut,obu,ogm,ogv,pmp,psp,qt,rcv,rm,rmvb,roq,rsd,smk,str,swf,thp,tod,tp,tr,ts,ty,\
ty+,usm,vc1,vc2,viv,vob,vro,vvc,vw,webm,wmv,wtv,xl,xmv,y4m,yop";

/// Read one entry out of the configured list into the lowercase form the lookups
/// use.
///
/// Every name in this list is a bare extension — unlike the archive list, which has
/// to carry the dotted `tar.gz` — so anything that is not one is dropped rather than
/// matched against.
pub fn sanitize_video_extensions(list: &str) -> Vec<String> {
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

/// Whether the configured list claims `path` as a video.
///
/// `.ts` and `.mts` are claimed by the text lists too, so a file of either name is
/// settled by its content: an MPEG-TS sync byte makes it the transport stream the
/// list says it is, and a file without one falls through to the text preview of the
/// TypeScript source it is.
pub fn matches_video_list(path: &Path, extensions: &[String]) -> bool {
    let Some(extension) = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase())
    else {
        return false;
    };

    if !extensions.contains(&extension) {
        return false;
    }

    if TYPESCRIPT_SHARED_EXTENSIONS.contains(&extension.as_str()) {
        return looks_like_mpegts(path);
    }

    true
}

/// Whether the configured list claims `path`, without asking whether video previews
/// are switched on.
pub fn is_video_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_video_list(path, &config.video_extensions))
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
