//! The viewer's own loading report, read off the pipe it writes it into: **the word a model's
//! picture is gated on**.
//!
//! **Why a pipe at all.** f3d draws a scene, and a scene's first render is not its last: VTK clears
//! the surface to the backdrop colour before the model is in it, so a source whose pixels have lit is
//! already true for the grey rectangle — and that is the black, then grey, then model a cold hover was
//! seen as (see `pin_film_thumbnail::the_source_drew`, whose gate this answers instead). What the
//! viewer does say, and says honestly, is when the file is loaded: measured on f3d 4.0.1, its standard
//! output carries
//! `\rLoading probe.stl :   0% |…| [00:00/00:00]\rLoading probe.stl : 100% |████…| [00:00/00:00]\r`,
//! flushed immediately and readable through a pipe while the process is still alive (150 bytes at
//! about 150-200 ms from a cold start). The `100%` is the semantic marker; the percentages before it
//! are the same line being written over.
//!
//! **The stream is piped only on a viewer's launch and drained only there** (see
//! `f3d_player::launch`, and `f3d_render::launch_args`, which asks for the report). The thread is the
//! pipe's only reader, which is what says a viewer can never be blocked by its own writing; nothing
//! here runs on an app thread, no read blocks one, and the thread ends when the pipe closes with the
//! process.
//!
//! **Two facts are kept per process, and no more**: whether the viewer has reported anything at all —
//! which says whether waiting for a marker is waiting for something that will never come — and when
//! the marker arrived. Both lifetime questions the gate asks are answered from the record (see
//! `pin_film_thumbnail::a_models_loading`).

use std::collections::HashMap;
use std::io::Read;
use std::sync::Mutex;
use std::time::Instant;

use once_cell::sync::Lazy;

/// The marker that says the file is loaded: the last percentage of the loading line. The lines are
/// separated by `\r` rather than `\n`, so nothing here splits them — the marker is one literal token
/// and a percentage short of it is a different token (see `the_file_is_loaded`).
const THE_FILE_IS_LOADED: &[u8] = b"100%";

/// How much of one read is kept for the next one: a marker split between two reads is put back
/// together by scanning the tail with the bytes that follow it, and nothing longer than the marker
/// itself can be half of one.
const LONGEST_TAIL: usize = THE_FILE_IS_LOADED.len() - 1;

/// How long one read may take: a chunk large enough that a loading line arrives whole, and small
/// enough that a report is one read rather than a buffer held for nothing.
const A_READ: usize = 512;

/// How many viewers' reports are kept at once. A model preview has one or two viewers alive — a
/// swap begins the second while the first is still the picture — so this bound is never reached by
/// live ones: what it drops is the record of a viewer this app ended long ago, whose pid nothing can
/// ask about any more (see `note_the_launch`).
const LIVE_REPORTS: usize = 4;

/// What one viewer's report has said so far, keyed by the process it is the report of.
#[derive(Clone, Copy)]
pub(super) struct ModelLoad {
    /// Whether the viewer has written anything at all. A viewer that says nothing is one whose
    /// loading this road cannot wait for (see `a_models_loading`).
    pub(super) reported: bool,
    /// When the marker arrived, once it has.
    pub(super) loaded: Option<Instant>,
    /// Whether the pipe has closed, which is the process ending or its report finishing: nothing
    /// more is coming either way.
    pub(super) ended: bool,
    /// When the report was begun, which is the clock a silent viewer is waited out on.
    pub(super) begun: Instant,
}

static LOADING: Lazy<Mutex<HashMap<u32, ModelLoad>>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// Begin the record of a launch, **before its window exists**: the gate asks this process's report
/// from the moment its source is adopted, and a source adopted before the record existed would be
/// answered as a viewer that reports nothing (see `the_load_of`).
///
/// A record is written whole rather than amended, so a pid Windows has handed out again — a viewer
/// this app ended, and the next one begun — is the new viewer's report and never the old one's.
pub(super) fn note_the_launch(pid: u32) {
    if pid == 0 {
        return;
    }
    let Ok(mut held) = LOADING.lock() else {
        return;
    };
    if held.len() >= LIVE_REPORTS {
        let oldest = held
            .iter()
            .min_by_key(|(_, report)| report.begun)
            .map(|(pid, _)| *pid);
        if let Some(oldest) = oldest {
            held.remove(&oldest);
        }
    }
    held.insert(
        pid,
        ModelLoad {
            reported: false,
            loaded: None,
            ended: false,
            begun: Instant::now(),
        },
    );
}

/// What this process's viewer has said about its loading so far, and nothing for a process that is
/// not one of this app's model viewers.
pub(super) fn the_load_of(pid: u32) -> Option<ModelLoad> {
    LOADING.lock().ok().and_then(|held| held.get(&pid).copied())
}

/// Read a viewer's report to its end on a thread of its own, and mark the record ended with it (see
/// the module note).
///
/// **The stream is any reader**, because the pipe a viewer's report arrives on is made by the launch
/// that begins it rather than by a `Command` (see `never_shown`): what this needs of it is `Read`,
/// to be the only one of it, and to outlive this call by as long as the thread does.
pub(super) fn drain_it(report: impl Read + Send + 'static, pid: u32) {
    let _ = std::thread::Builder::new()
        .name("model-load".to_string())
        .spawn(move || {
            read_the_report(report, pid);
            note_the_end(pid);
        });
}

/// The end of a report: nothing that did not arrive by now is coming.
fn note_the_end(pid: u32) {
    if let Ok(mut held) = LOADING.lock() {
        if let Some(report) = held.get_mut(&pid) {
            report.ended = true;
        }
    }
}

/// Whether a window of a viewer's output carries the marker: `100%` where it lies, which is the last
/// percentage of a loading line and nothing else on that line.
///
/// **No line is split off**, and that is deliberate: the report separates its lines with `\r`, and a
/// `100%` is one literal token wherever it stands in the stream — splitting on `\r` first would also
/// throw away the half of a token a read happens to end in (see `read_the_report`).
fn the_file_is_loaded(output: &[u8]) -> bool {
    output
        .windows(THE_FILE_IS_LOADED.len())
        .any(|window| window == THE_FILE_IS_LOADED)
}

/// Read a viewer's report in chunks until the pipe closes, writing the two facts it is read for.
///
/// **A marker split between two reads is still a marker**: the last few bytes of every read are read
/// again beside the bytes that follow them, and a token longer than that tail cannot be half-written
/// into it.
fn read_the_report(mut report: impl Read, pid: u32) {
    let mut tail: Vec<u8> = Vec::with_capacity(LONGEST_TAIL);
    let mut chunk = [0u8; A_READ];
    loop {
        let count = match report.read(&mut chunk) {
            Ok(0) | Err(_) => return,
            Ok(count) => count,
        };
        let mut seen = std::mem::take(&mut tail);
        seen.extend_from_slice(&chunk[..count]);
        note_the_bytes(pid, &seen);
        tail = seen[seen.len().saturating_sub(LONGEST_TAIL)..].to_vec();
    }
}

/// Write what one window of the report says: that the viewer reports at all, and — once, and never
/// overwritten — when its file was loaded.
fn note_the_bytes(pid: u32, seen: &[u8]) {
    let loaded = the_file_is_loaded(seen);
    let Ok(mut held) = LOADING.lock() else {
        return;
    };
    let Some(report) = held.get_mut(&pid) else {
        return;
    };
    report.reported |= !seen.is_empty();
    if loaded && report.loaded.is_none() {
        report.loaded = Some(Instant::now());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A report that arrives in pieces, the way a pipe hands one over: nothing about a marker may
    /// depend on where a read happened to end (see `read_the_report`).
    struct InPieces<'a> {
        bytes: &'a [u8],
        each: usize,
    }

    impl Read for InPieces<'_> {
        fn read(&mut self, into: &mut [u8]) -> std::io::Result<usize> {
            if self.bytes.is_empty() {
                return Ok(0);
            }
            let count = self.each.min(self.bytes.len());
            into[..count].copy_from_slice(&self.bytes[..count]);
            self.bytes = &self.bytes[count..];
            Ok(count)
        }
    }

    /// The pids this module's own tests report under: no player's and no test's, so a record left
    /// behind by one of them cannot be read as another's (see `note_the_launch`).
    const A_VIEWER: u32 = 0xF3D0_0001;
    const ANOTHER_VIEWER: u32 = 0xF3D0_0002;

    /// **The marker is the file loaded, and a percentage short of it is not the marker.** The report's
    /// lines are separated by `\r` and the `0%` of a line being written over must never be read as the
    /// whole of the load.
    #[test]
    fn only_the_last_percentage_is_the_file_loaded() {
        assert!(!the_file_is_loaded(b""));
        assert!(!the_file_is_loaded(
            b"\rLoading probe.stl :   0% |  | [00:00/00:00]"
        ));
        assert!(!the_file_is_loaded(
            b"\rLoading probe.stl :  10% |# | [00:00/00:00]"
        ));
        assert!(
            !the_file_is_loaded(b"\rLoading probe.stl : 100"),
            "a token whose last byte has not arrived yet is not the marker"
        );
        assert!(
            !the_file_is_loaded(b"\rLoading probe.stl : 10"),
            "and neither is one still being written into"
        );
        assert!(the_file_is_loaded(
            b"\rLoading probe.stl :   0% |  | [00:00/00:00]\rLoading probe.stl : 100% |####| \
              [00:00/00:00]\r"
        ));
    }

    /// **A marker the pipe splits is still the marker**, whatever byte a read ends on, and a viewer
    /// that writes anything at all is one this road can wait for.
    #[test]
    fn a_report_is_read_whole_however_the_pipe_splits_it() {
        let report = b"\rLoading probe.stl :   0% |  | [00:00/00:00]\rLoading probe.stl : 100% \
                       |####| [00:00/00:00]\r";

        for each in [1, 2, 3, 5, 7, 64] {
            note_the_launch(A_VIEWER);
            read_the_report(
                InPieces {
                    bytes: report,
                    each,
                },
                A_VIEWER,
            );
            let read = the_load_of(A_VIEWER).expect("the record the launch made");
            assert!(
                read.reported,
                "a viewer that wrote is one that reports, read {each} bytes at a time"
            );
            assert!(
                read.loaded.is_some(),
                "and the marker is found wherever a read ends, read {each} bytes at a time"
            );
            assert!(
                !read.ended,
                "the end is noted by the thread, not by the reading"
            );
        }
    }

    /// **A viewer that wrote nothing is one the record says nothing about**, and the close of its pipe
    /// is the end of the wait rather than a report that is still coming.
    #[test]
    fn a_silent_viewer_is_ended_and_never_loaded() {
        note_the_launch(ANOTHER_VIEWER);
        read_the_report(
            InPieces {
                bytes: b"",
                each: 1,
            },
            ANOTHER_VIEWER,
        );
        note_the_end(ANOTHER_VIEWER);

        let read = the_load_of(ANOTHER_VIEWER).expect("the record the launch made");
        assert!(!read.reported, "nothing was said");
        assert_eq!(read.loaded, None, "and nothing was loaded");
        assert!(read.ended, "and its pipe is closed");

        assert!(
            the_load_of(0xF3D0_0003).is_none(),
            "a process that was never begun is no viewer of this app's"
        );
    }
}
