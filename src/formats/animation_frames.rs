//! The pace of a picture that plays: how many frames it holds, where each one begins,
//! and how long the whole of them lasts.
//!
//! A seek bar for an animation ends at the file's own last frame and lands on a frame
//! rather than a second of video, so what it needs is the run of frames and the time
//! each one is held. Both are in the container's own block list — a GIF's graphic
//! control extensions, an APNG's `fcTL` chunks, an animated WebP's `ANMF` headers — and
//! the walks below read those alone: no pixel is decoded to answer any of it, the same
//! bargain `head` makes for the two questions it asks of the same blocks.
//!
//! What is read is the file whole, under the same budget every decoder that reads
//! bytes rather a reader is held to, and the answer is held under the file's own
//! version so a bar asked for it on every repaint costs one read per file.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

use once_cell::sync::Lazy;

use crate::config::config::read_within_budget;

/// The floor under a frame's own delay, which is the floor the decoders themselves
/// apply (`MIN_ANIMATION_FRAME_DELAY_MS`): a file saying zero — or a still timebase,
/// which is zero too — would otherwise be a playhead that runs through the whole
/// animation in the few milliseconds between two repaints.
const FRAME_DELAY_FLOOR_MS: u32 = 16;

/// How many timelines are held. What it bounds is a pointer dragged across a folder of
/// animations, and it is the same figure the heads are held under.
const TIMELINES_MAX_ENTRIES: usize = 256;

/// One playing picture's frames in the order they play, and the time each one holds.
pub(crate) struct FrameTimeline {
    /// Where each frame begins, in milliseconds from the file's first frame, and the
    /// whole of them at the end: one longer than `delays_ms`.
    starts_ms: Vec<u64>,
    delays_ms: Vec<u32>,
}

impl FrameTimeline {
    /// A timeline of the delays given, for a test that asks what a run of them means rather than
    /// what a container holds (see `FrameTimeline::new`).
    #[cfg(test)]
    pub(crate) fn of_delays(delays_ms: &[u32]) -> Self {
        Self::new(delays_ms.to_vec())
    }

    /// A timeline of `delays_ms`, each floored at the delay floor the decoders use.
    fn new(delays_ms: Vec<u32>) -> Self {
        let mut starts_ms = Vec::with_capacity(delays_ms.len() + 1);
        let mut at = 0u64;
        starts_ms.push(0);
        for delay in &delays_ms {
            at = at.saturating_add(u64::from((*delay).max(FRAME_DELAY_FLOOR_MS)));
            starts_ms.push(at);
        }

        Self {
            starts_ms,
            delays_ms,
        }
    }

    /// How many frames the file plays.
    pub(crate) fn frames(&self) -> usize {
        self.delays_ms.len()
    }

    /// How long the file lasts when played at its own pace, in seconds.
    pub(crate) fn seconds(&self) -> f64 {
        self.starts_ms.last().copied().unwrap_or(0) as f64 / 1000.0
    }

    /// The pace the file itself keeps, in frames a second — what the FPS slider opens at.
    ///
    /// Nothing is rounded here: the answer is a record of the file's own delays, and the
    /// one thing that has a bound on it (the slider's top) is the control's question, not
    /// this one's (see `current_animated_fps_max`).
    pub(crate) fn rate(&self) -> f64 {
        let seconds = self.seconds();
        if seconds <= 0.0 || self.frames() == 0 {
            return 1.0;
        }

        self.frames() as f64 / seconds
    }

    /// Where the frame at `index` begins, in seconds from the file's first frame.
    pub(crate) fn begins_at(&self, index: usize) -> f64 {
        self.starts_ms
            .get(index)
            .or_else(|| self.starts_ms.last())
            .copied()
            .unwrap_or(0) as f64
            / 1000.0
    }

    /// The frame a bar dragged to `share` of its length lands on: the end of the bar is the end of
    /// the file, which is the last frame's own run and so the last frame.
    ///
    /// It is time rather than count, because the frames of a file are not all held for
    /// the same stretch — a 30-frame animation whose last frame waits two seconds is not
    /// one where frame fifteen is halfway through it.
    pub(crate) fn index_at_share(&self, share: f64) -> usize {
        let share = share.clamp(0.0, 1.0);
        let wanted = share * self.seconds();

        self.index_at_seconds(wanted)
    }

    /// The frame a playhead at `seconds` is on, being the last frame whose own run the
    /// second falls inside — and the last frame of all for a second at or past the end.
    pub(crate) fn index_at_seconds(&self, seconds: f64) -> usize {
        if self.frames() == 0 {
            return 0;
        }

        let wanted = (seconds.max(0.0) * 1000.0).round() as u64;
        self.starts_ms
            .partition_point(|start| *start <= wanted)
            .saturating_sub(1)
            .min(self.frames() - 1)
    }
}

/// The key a timeline is held under: the file, the version of it, and how long it was.
#[derive(Clone, PartialEq, Eq, Hash)]
struct TimelineKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

static TIMELINES: Lazy<Mutex<HashMap<TimelineKey, Option<Arc<FrameTimeline>>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// The timeline of the file at `path`, or nothing where its container says nothing
/// that can be walked — which is a bar with no length to drag, rather an animation the
/// app will not play.
///
/// Answers are held, including the ones that answer nothing: a file walked once and
/// found unreadable is a file not walked again on every repaint of the same bar.
pub(crate) fn timeline(path: &Path) -> Option<Arc<FrameTimeline>> {
    let key = key_of(path);
    if let Ok(timelines) = TIMELINES.lock() {
        if let Some(held) = timelines.get(&key) {
            return held.clone();
        }
    }

    let answer = read_timeline(path).map(Arc::new);
    if let Ok(mut timelines) = TIMELINES.lock() {
        if timelines.len() >= TIMELINES_MAX_ENTRIES {
            timelines.clear();
        }
        timelines.insert(key, answer.clone());
    }

    answer
}

fn key_of(path: &Path) -> TimelineKey {
    let metadata = std::fs::metadata(path);
    let len = metadata.as_ref().map(|meta| meta.len()).unwrap_or(0);
    let modified = metadata.ok().and_then(|meta| meta.modified().ok());

    TimelineKey {
        path: path.to_path_buf(),
        modified,
        len,
    }
}

/// The file's own delays, walked out of its container's block list.
fn read_timeline(path: &Path) -> Option<FrameTimeline> {
    let bytes = read_within_budget(path)?;

    let delays = if bytes.starts_with(b"GIF8") {
        gif_delays(&bytes)?
    } else if bytes.len() >= 12 && &bytes[..4] == b"RIFF" && &bytes[8..12] == b"WEBP" {
        webp_delays(&bytes)?
    } else if bytes.starts_with(&PNG_SIGNATURE) {
        apng_delays(&bytes)?
    } else {
        return None;
    };

    if delays.is_empty() {
        return None;
    }

    Some(FrameTimeline::new(delays))
}

const PNG_SIGNATURE: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];

/// A GIF's frames' delays, in the order they play: one block walk, a graphic control
/// extension's own delay read for the image descriptor that follows it.
///
/// Every other block is stepped over by its own length — the walk is the same one
/// `head` makes to count frames, with the delay kept rather the count alone — and a
/// file that ends before its trailer yields the frames it did declare.
fn gif_delays(bytes: &[u8]) -> Option<Vec<u32>> {
    if bytes.len() < 13 {
        return None;
    }

    let packed = bytes[10];
    let mut at = 13usize;
    if packed & 0x80 != 0 {
        at = at.checked_add(colour_table_bytes(packed))?;
    }

    let mut delays = Vec::new();
    let mut pending = 0u32;

    while at < bytes.len() {
        match bytes[at] {
            // An image descriptor: the frame this extension's delay belongs to.
            0x2C => {
                let descriptor = bytes.get(at + 1..at + 10)?;
                at += 10;
                if descriptor[8] & 0x80 != 0 {
                    at = at.checked_add(colour_table_bytes(descriptor[8]))?;
                }
                // The LZW minimum code size, then the pixels as sub-blocks.
                at = at.checked_add(1)?;
                at = skip_sub_blocks(bytes, at)?;
                delays.push(pending);
                pending = 0;
            }
            // An extension: its label, then blocks that carry their own lengths.
            0x21 => {
                let label = bytes.get(at + 1).copied()?;
                at += 2;
                if label == 0xF9 {
                    let body = bytes.get(at..at + 5)?;
                    if body[0] == 4 {
                        // Hundredths of a second, which is the GIF's own timebase.
                        pending = u32::from(u16::from_le_bytes([body[2], body[3]])) * 10;
                    }
                }
                at = skip_sub_blocks(bytes, at)?;
            }
            // The trailer: the file ends here, and what was read is the whole of it.
            0x3B => break,
            _ => break,
        }
    }

    Some(delays)
}

/// The bytes a GIF's colour table takes, from the packed byte that says whose it is.
fn colour_table_bytes(packed: u8) -> usize {
    3 * (1usize << (u32::from(packed & 0x07) + 1))
}

/// Step over one run of GIF sub-blocks: a length byte per block, ending at a
/// zero-length one.
fn skip_sub_blocks(bytes: &[u8], mut at: usize) -> Option<usize> {
    loop {
        let length = usize::from(*bytes.get(at)?);
        at += 1;
        if length == 0 {
            return Some(at);
        }
        at = at.checked_add(length)?;
    }
}

/// An animated WebP's frame durations, from the `ANMF` chunk headers of its `ANIM`
/// extension: each carries its own duration in milliseconds, and the chunks between
/// them are stepped over by their declared lengths.
fn webp_delays(bytes: &[u8]) -> Option<Vec<u32>> {
    let mut at = 12usize;
    let mut delays = Vec::new();

    while at + 8 <= bytes.len() {
        let kind = &bytes[at..at + 4];
        let length =
            u32::from_le_bytes([bytes[at + 4], bytes[at + 5], bytes[at + 6], bytes[at + 7]])
                as usize;
        let body = at + 8;

        if kind == b"ANMF" {
            // A frame's own header: where it lands, how big it is, and then the
            // duration as a 24-bit little-endian count of milliseconds.
            let header = bytes.get(body..body.checked_add(16)?)?;
            let duration = u32::from_le_bytes([header[12], header[13], header[14], 0]);
            delays.push(duration);
        }

        // Chunks are padded to an even length.
        at = body.checked_add(length)?.checked_add(length % 2)?;
    }

    Some(delays)
}

/// An APNG's frame delays, read from the `fcTL` chunks that carry them — each a
/// delay numerator and denominator, with a denominator of zero standing for the
/// hundredths a GIF would say.
fn apng_delays(bytes: &[u8]) -> Option<Vec<u32>> {
    let mut at = 8usize;
    let mut delays = Vec::new();

    while at + 8 <= bytes.len() {
        let length =
            u32::from_be_bytes([bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]]) as usize;
        let kind = &bytes[at + 4..at + 8];
        let body = at + 8;

        if kind == b"fcTL" {
            // The sequence number, the frame's box, and then the delay: four-byte
            // fields ahead of it, and the numerator and denominator behind them.
            let header = bytes.get(body..body.checked_add(24)?)?;
            let numerator = u32::from(u16::from_be_bytes([header[20], header[21]]));
            let denominator = match u16::from_be_bytes([header[22], header[23]]) {
                0 => 100,
                given => u32::from(given),
            };
            delays.push(numerator.saturating_mul(1000) / denominator);
        }

        if kind == b"IEND" {
            break;
        }

        // The chunk's body, then its CRC.
        at = body.checked_add(length)?.checked_add(4)?;
    }

    Some(delays)
}

#[cfg(test)]
mod tests;
