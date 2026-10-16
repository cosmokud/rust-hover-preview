//! The length of a raw AAC stream: the frames the file holds, counted.
//!
//! An `.aac` file is a bare stream — an ADTS frame after an ADTS frame, and nothing in front of
//! them that says how long the whole is — so the length of one is not a thing either engine
//! this app plays it with can read; it is a thing they *estimate* from the stream's bit rate,
//! and FFmpeg says as much while it does it: "Estimating duration from bitrate, this may be
//! inaccurate". Inaccurate is what it is, in both directions and by more than a little: a
//! twenty-four second stream written here was read as twenty-six seconds by FFmpeg and as
//! twenty-five by the engine Windows has, and an hour of one was read as sixty-six minutes.
//! What a sound's card draws its bar and its clock against is that number, so a stream the
//! estimate reads short is a card saying the end of the file has been reached — `2:40 / 2:40`,
//! with the bar filled to its right-hand end — while the last seconds of the sound are still
//! being heard (see `audio_clock` and `fill_span`).
//!
//! What a raw stream does say exactly is how many frames it holds, and what a frame holds is
//! exactly 1024 samples of the rate its own header names: the length is the frames times that.
//! The frames are counted by walking them, each header giving the length of the frame it
//! begins, which costs one sequential read of the file and nothing else — no decoder, no
//! process, and no second look at a file that is not a raw stream at all, since a raw stream
//! begins with a frame and a container begins with a box of its own (`ftyp`, for one). A minute
//! of sound is about a megabyte of file, so what the walk reads is the file's own size and the
//! walk itself is nothing.
//!
//! The count is only used where it is beyond doubt, because the estimate it replaces is only
//! ever wrong about the length while a count that is wrong is a *new* wrong length: a walk that
//! had to step over a byte that is not a frame, that ran out before the frame a header promised,
//! or whose frames disagree about the rate they are written at — the one field that has to agree,
//! since it is what a frame's 1024 samples are counted in, where a channel layout a stream
//! changes its mind about halfway through is nothing a length is counted by — is a file this
//! answers nothing for. The length the probe read stands for one of those, which is what every
//! file that is not a raw AAC stream gets anyway (see [`duration`]).

use std::fs::File;
use std::io::{BufRead, BufReader, ErrorKind, Read, Seek, SeekFrom};
use std::path::Path;

/// The rates an ADTS header names by the index its four bits carry. Thirteen and fourteen are
/// reserved and fifteen says the rate is spelled out in the frame's own payload, which nothing
/// this app is handed writes: a header naming one of those is not one a length is counted from
/// (see [`frame_of`]).
const RATES: [u32; 13] = [
    96_000, 88_200, 64_000, 48_000, 44_100, 32_000, 24_000, 22_050, 16_000, 12_000, 11_025, 8_000,
    7_350,
];

/// The samples one AAC frame holds, in the units of the rate its header names — and the header's
/// own two bits for raw data blocks say how many of those the frame carries: one for every
/// stream but the few that pack several blocks into one frame (see [`frame_of`]).
const SAMPLES_PER_BLOCK: u64 = 1024;

/// The buffer a file is walked through, which is as much as the longest a frame can be — the
/// thirteen bits its header gives its own length say no more than this — so that the whole of
/// the frame whose header stands at the front of the buffer is inside it.
const READ_BYTES: usize = 8192;

/// One frame's header, read for the two things it decides: where the next frame begins and how
/// much time this one holds.
struct Frame {
    /// The whole of the frame, counted from its own front: its header, its payload, and the two
    /// bytes of CRC a frame with one carries — a length is a length.
    length: usize,
    /// The samples a second the stream is written at, off the header and not off the probe: it
    /// is the rate a frame's samples are counted in, which is the rate the stream is *coded* at
    /// and, for a stream with SBR, half of the rate a player hands over beside it.
    rate: u32,
    /// The samples this frame holds.
    samples: u64,
}

/// How long the raw AAC stream in `path` plays, where the whole of the file is ADTS frames —
/// and nothing where the file is anything else, which is the answer every container, every
/// other sound and every damaged stream is given (see the module's note).
pub fn duration(path: &Path) -> Option<f64> {
    let mut file = BufReader::with_capacity(READ_BYTES, File::open(path).ok()?);
    skip_id3(&mut file);

    // A raw stream begins with a frame, and this is where a file that is not one is turned away
    // before it has been read at all: a container's box, another sound's frames, a picture's
    // head — none of them is the front of a frame. Turning it away here is also what keeps the
    // walk from reading an `.m4a`, whose frames are thousands of bytes inside a container of
    // boxes, to its end in search of one.
    let mut header = [0u8; 7];
    if !read_or_end(&mut file, &mut header)? {
        return None;
    }
    frame_of(&header)?;

    let mut body = Vec::new();
    let mut frames = 0u64;
    let mut samples = 0u64;
    let mut rate = None;
    // Whether every byte of the file was frames: a byte that had to be stepped over, a frame
    // that ran out before its own length said it would, and a header left half-read are all a
    // file nothing is counted for (see the module's note).
    let mut whole = true;

    loop {
        let Some(frame) = frame_of(&header) else {
            // A byte that is not the front of a frame: a stream damaged where this stands, or
            // something else entirely. The walk steps over it rather than giving up, because the
            // frames after it are still frames and counting them costs a walk either way — and
            // stepping over any byte at all is itself the end of the answer.
            whole = false;
            header.copy_within(1.., 0);
            if !read_or_end(&mut file, &mut header[6..])? {
                break;
            }
            continue;
        };

        // Every frame of a stream is written at one rate, and a file whose frames disagree is
        // not one stream.
        match rate {
            Some(rate) => whole &= rate == frame.rate,
            None => rate = Some(frame.rate),
        }

        // The frame is read past rather than kept: all the walk wants of a file is how many
        // frames there are and how much time each of them carries.
        body.resize(frame.length - header.len(), 0);
        if !read_or_end(&mut file, &mut body)? {
            whole = false;
            break;
        }

        frames += 1;
        samples += frame.samples;

        // The end of the file is the end of the walk: nothing at all left to read. A few bytes
        // that cannot make a whole frame are a stream cut off, and a read that fails outright is
        // no better — neither is a file this answers for (see `whole`).
        match file.fill_buf() {
            Ok([]) => break,
            Ok(_) => {}
            Err(_) => {
                whole = false;
                break;
            }
        }

        if !read_or_end(&mut file, &mut header)? {
            whole = false;
            break;
        }
    }

    let rate = rate?;
    let seconds = samples as f64 / f64::from(rate);

    (whole && frames > 0 && seconds > 0.0).then_some(seconds)
}

/// The frame a header begins, where these seven bytes are the front of one.
///
/// Nothing here is guessed at: twelve bits of sync, a layer of zero beneath them, an object type
/// that is one of the three the format defines, a rate its own index names, and a length long
/// enough to hold the header — anything else is a byte that is not the front of a frame, which
/// the caller either refuses the file over or steps over and refuses it later for.
fn frame_of(header: &[u8; 7]) -> Option<Frame> {
    // The sync's twelve bits — `0xFF`, and the top half of the byte after it — with a layer of
    // zero beneath them. The version bit between the two (MPEG-2 against MPEG-4) and the low bit
    // of the second byte, which says whether the frame carries a CRC, are masked out rather than
    // compared: neither is anything a length is counted by.
    if header[0] != 0xFF || header[1] & 0xF6 != 0xF0 {
        return None;
    }

    // The two bits of object type, of which the fourth names nothing any encoder writes. The
    // three it can name hold 1024 samples a frame, which is what the samples below are counted
    // in.
    if header[2] & 0xC0 == 0xC0 {
        return None;
    }

    let rate = RATES.get(usize::from((header[2] >> 2) & 0x0F)).copied()?;

    // The frame's own length, thirteen bits: the header's seven bytes and its payload, where the
    // two CRC bytes of a frame with one sit among them. A length shorter than the header it is
    // written in is a length that would step backwards, which is not a frame.
    let length = ((usize::from(header[3]) & 0x03) << 11)
        | (usize::from(header[4]) << 3)
        | (usize::from(header[5]) >> 5);
    if length < header.len() {
        return None;
    }

    // The last two bits are the frame's raw data blocks, less one: a frame of two blocks holds
    // two frames' worth of samples, and its own length says so as well.
    let samples = SAMPLES_PER_BLOCK * (u64::from(header[6] & 0x03) + 1);

    Some(Frame {
        length,
        rate,
        samples,
    })
}

/// Read the whole of `into`, answering whether all of it was there: the end of the file is a
/// `false`, which is the caller's business rather than a fault, and anything else that goes
/// wrong is a question this gives up on.
fn read_or_end(file: &mut BufReader<File>, into: &mut [u8]) -> Option<bool> {
    match file.read_exact(into) {
        Ok(()) => Some(true),
        Err(error) if error.kind() == ErrorKind::UnexpectedEof => Some(false),
        Err(_) => None,
    }
}

/// Step over the ID3v2 tag a file may begin with, which is not part of the stream behind it: a
/// player skips one the same way, and a walk that did not would answer nothing for every `.aac`
/// an encoder had written a tag onto.
///
/// What a tag is, is ten bytes of header and then the body: its size is written in the header as
/// four seven-bit numbers that leave the top bit of each alone, and ten more bytes follow the
/// body where the header's flags say the tag ends with a footer of its own. Nothing is stepped
/// over where the file does not begin with a tag, which is also where nothing is read from
/// beyond those ten bytes.
fn skip_id3(file: &mut BufReader<File>) {
    let Ok(at) = file.stream_position() else {
        return;
    };

    let mut head = [0u8; 10];
    if file.read_exact(&mut head).is_err() || &head[0..3] != b"ID3" {
        // Not a tag, or not ten bytes of one: nothing was walked over, and the walk goes on from
        // the front of the file.
        let _ = file.seek(SeekFrom::Start(at));
        return;
    }

    let size = head[6..10]
        .iter()
        .fold(0u64, |size, byte| (size << 7) | u64::from(byte & 0x7F));
    let footer = if head[5] & 0x10 == 0 { 0 } else { 10 };

    let _ = file.seek(SeekFrom::Start(at + 10 + size + footer));
}

/// One ADTS frame laid out where the format says its fields sit: a header of seven bytes, a
/// payload of nothing in particular, and `blocks` raw data blocks carried in it — a frame of
/// more than one block being a thing a few streams write and a length counted the same either
/// way (see [`frame_of`]). What the tests of the walk and of the length a probe is held at are
/// written from, since both need a file that is exactly what a raw AAC stream is.
#[cfg(test)]
pub(crate) fn a_frame_holding(rate_index: u8, payload: usize, blocks: u8) -> Vec<u8> {
    let length = 7 + payload;

    let mut frame = vec![
        0xFF,
        0xF1,                     // the sync, MPEG-4, a layer of zero, and no CRC
        0x40 | (rate_index << 2), // LC, the rate, and mono
        ((length >> 11) as u8) & 0x03,
        (length >> 3) as u8,
        (((length & 0x07) as u8) << 5) | 0x1F,
        0xFC | (blocks - 1),
    ];
    frame.resize(length, 0x5A);

    frame
}

/// The same, one raw data block to the frame, which is what every ordinary stream writes.
#[cfg(test)]
pub(crate) fn a_frame(rate_index: u8, payload: usize) -> Vec<u8> {
    a_frame_holding(rate_index, payload, 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A file for one test of these, written where every other test of this app's writes one.
    fn a_file(case: &str, bytes: &[u8]) -> std::path::PathBuf {
        let folder = std::env::temp_dir().join("rust-hover-preview-aac-stream");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let path = folder.join(case);
        std::fs::write(&path, bytes).expect("a written file");

        path
    }

    /// An ID3v2 tag of `body` bytes of something, and then what follows it.
    fn a_tag(body: usize) -> Vec<u8> {
        let size = body as u32;
        let mut tag = vec![
            b'I',
            b'D',
            b'3',
            0x04,
            0x00, // the version, which nothing reads
            0x00, // the flags, which say there is no footer
            ((size >> 21) & 0x7F) as u8,
            ((size >> 14) & 0x7F) as u8,
            ((size >> 7) & 0x7F) as u8,
            (size & 0x7F) as u8,
        ];
        tag.resize(10 + body, 0x00);

        tag
    }

    /// A stream's length is its frames counted and the time each of them holds: 1024 samples of
    /// the rate the file's own headers name, whatever the frames' sizes are.
    #[test]
    fn a_streams_length_is_the_frames_it_holds_counted() {
        let mut bytes = Vec::new();
        for payload in [10, 300, 7, 40] {
            bytes.extend(a_frame(4, payload));
        }
        let path = a_file("four-frames.aac", &bytes);

        assert_eq!(
            duration(&path),
            Some(4.0 * 1024.0 / 44_100.0),
            "four frames of forty-four one are four times 1024 samples of it"
        );

        // And a frame that holds two of the format's blocks holds twice the samples, which is
        // what its own header's last two bits are for.
        let past = a_file("blocks.aac", &a_frame_holding(4, 200, 2));
        assert_eq!(
            duration(&past),
            Some(2.0 * 1024.0 / 44_100.0),
            "a frame of two blocks is two frames' worth of samples"
        );
    }

    /// A tag in front of the frames is not part of the stream: it is stepped over, and its own
    /// size is what says how far — a tag walked over wrongly would leave the walk in its body,
    /// where nothing is a frame.
    #[test]
    fn a_tag_in_front_of_a_stream_is_not_part_of_it() {
        let mut bytes = a_tag(1_000);
        for _ in 0..3 {
            bytes.extend(a_frame(3, 120)); // 48000
        }
        let path = a_file("tagged.aac", &bytes);

        assert_eq!(
            duration(&path),
            Some(3.0 * 1024.0 / 48_000.0),
            "the frames behind the tag are counted and the tag is not"
        );
    }

    /// A file that is not a raw stream is answered with nothing, and answered so before it has
    /// been read: a container begins with its own box, and a sound that begins anywhere but at
    /// a frame is not a stream this side can count the frames of.
    #[test]
    fn a_file_that_is_not_a_stream_of_frames_answers_nothing() {
        let container = a_file("container.m4a", b"\0\0\0\x18ftypM4A \0\0\0\0M4A mp42isom");
        assert_eq!(
            duration(&container),
            None,
            "a container's box is not the front of a frame"
        );

        let mut mid_stream = a_frame(4, 40);
        mid_stream.drain(0..3);
        assert_eq!(
            duration(&a_file("cut-front.aac", &mid_stream)),
            None,
            "a stream whose front is missing is not one to count"
        );

        assert_eq!(
            duration(&a_file("nothing.aac", b"")),
            None,
            "and an empty file holds no frames at all"
        );
    }

    /// A stream is answered for only where every byte of it was frames: one byte that had to be
    /// stepped over, one frame that ran out before its own length said it would, or one frame
    /// written at a rate the others are not is a file whose count is not its length.
    #[test]
    fn a_stream_anything_had_to_be_stepped_over_in_answers_nothing() {
        let mut straddled = a_frame(4, 40);
        straddled.extend_from_slice(b"junk");
        straddled.extend(a_frame(4, 40)); // a byte of something between two frames
        assert_eq!(
            duration(&a_file("straddled.aac", &straddled)),
            None,
            "a byte that is not a frame is not one to count around"
        );

        let mut cut = a_frame(4, 40);
        cut.extend_from_slice(&a_frame(4, 500)[0..20]); // a frame that is not all there
        assert_eq!(
            duration(&a_file("cut-frame.aac", &cut)),
            None,
            "a frame the file ran out inside is not one that played"
        );

        let mut disagreeing = a_frame(4, 40);
        disagreeing.extend(a_frame(3, 40)); // 48000 against 44100
        assert_eq!(
            duration(&a_file("disagreeing.aac", &disagreeing)),
            None,
            "two rates are two streams, and this is one file"
        );
    }

    /// A frame whose header carries a CRC is counted like any other: its length covers the two
    /// bytes the CRC is written in, which is what the walk steps over a frame by.
    #[test]
    fn a_frame_with_a_crc_is_counted_like_any_other() {
        let mut frame = a_frame(4, 50);
        frame[1] = 0xF0; // the low bit down: a CRC, written inside the frame's own length
        let mut bytes = frame.clone();
        bytes.extend(a_frame(4, 50));

        assert_eq!(
            duration(&a_file("crc.aac", &bytes)),
            Some(2.0 * 1024.0 / 44_100.0),
            "two frames, one of them with a CRC, are two frames' worth of samples"
        );
    }
}
