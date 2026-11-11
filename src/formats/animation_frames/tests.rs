//! The walks the timeline is read by, on the smallest containers that can say what
//! they are — and on ones that say nothing at all.

use super::*;

/// A GIF header, a logical screen descriptor of `packed`, and the global colour table
/// the packed byte asks for.
fn gif_head(packed: u8) -> Vec<u8> {
    let mut bytes = b"GIF89a".to_vec();
    bytes.extend_from_slice(&[0x02, 0x00, 0x02, 0x00]);
    bytes.push(packed);
    bytes.extend_from_slice(&[0x00, 0x00]);
    if packed & 0x80 != 0 {
        bytes.extend_from_slice(&vec![0u8; colour_table_bytes(packed)]);
    }
    bytes
}

/// A GIF frame: its descriptor, the LZW minimum code size, and one empty sub-block run.
fn gif_frame(with_local_table: bool) -> Vec<u8> {
    let packed = if with_local_table { 0x81 } else { 0x00 };
    let mut bytes = vec![0x2C, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x02, 0x00, packed];
    if with_local_table {
        bytes.extend_from_slice(&vec![0u8; colour_table_bytes(packed)]);
    }
    bytes.extend_from_slice(&[0x02, 0x00]);
    bytes
}

/// A GIF graphic control extension holding `hundredths` hundredths of a second.
fn gif_delay(hundredths: u16) -> Vec<u8> {
    vec![
        0x21,
        0xF9,
        0x04,
        0x00,
        (hundredths & 0xFF) as u8,
        (hundredths >> 8) as u8,
        0x00,
        0x00,
    ]
}

/// A PNG signature and the `acTL` that makes it an animation of `frames`.
fn apng_head(frames: u32) -> Vec<u8> {
    let mut bytes = PNG_SIGNATURE.to_vec();
    bytes.extend_from_slice(&13u32.to_be_bytes());
    bytes.extend_from_slice(b"IHDR");
    bytes.extend_from_slice(&[0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02]);
    bytes.extend_from_slice(&[0x08, 0x06, 0x00, 0x00, 0x00]);
    bytes.extend_from_slice(&[0u8; 4]);
    bytes.extend_from_slice(&8u32.to_be_bytes());
    bytes.extend_from_slice(b"acTL");
    bytes.extend_from_slice(&frames.to_be_bytes());
    bytes.extend_from_slice(&[0u8; 4]);
    bytes.extend_from_slice(&[0u8; 4]);
    bytes
}

/// An APNG frame control chunk holding `numerator`/`denominator` of a second.
fn apng_frame_control(numerator: u16, denominator: u16) -> Vec<u8> {
    let mut body = vec![0u8; 20];
    body.extend_from_slice(&numerator.to_be_bytes());
    body.extend_from_slice(&denominator.to_be_bytes());
    // The dispose and blend operations that close the frame's own control chunk.
    body.extend_from_slice(&[0u8; 2]);
    let mut bytes = 26u32.to_be_bytes().to_vec();
    bytes.extend_from_slice(b"fcTL");
    bytes.extend_from_slice(&body);
    bytes.extend_from_slice(&[0u8; 4]);
    bytes
}

/// A RIFF WebP front, and the chunks behind it.
fn webp_with(chunks: &[(&[u8; 4], Vec<u8>)]) -> Vec<u8> {
    let mut body = b"WEBP".to_vec();
    for (kind, body_bytes) in chunks {
        body.extend_from_slice(&kind[..]);
        body.extend_from_slice(&(body_bytes.len() as u32).to_le_bytes());
        body.extend_from_slice(body_bytes);
        if body_bytes.len() % 2 != 0 {
            body.push(0);
        }
    }
    let mut bytes = b"RIFF".to_vec();
    bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
    bytes.extend_from_slice(&body);
    bytes
}

/// An `ANMF` frame chunk lasting `milliseconds`.
fn webp_frame(milliseconds: u32) -> Vec<u8> {
    let mut body = vec![0u8; 12];
    body.extend_from_slice(&milliseconds.to_le_bytes()[..3]);
    body.extend_from_slice(&[0u8; 4]);
    body
}

/// A timeline built from delays, which is what every walk below answers with.
fn timeline_of(delays: &[u32]) -> FrameTimeline {
    FrameTimeline::new(delays.to_vec())
}

#[test]
fn a_gif_reports_every_frame_and_the_delay_it_holds() {
    let mut bytes = gif_head(0x81);
    bytes.extend(gif_delay(10));
    bytes.extend(gif_frame(false));
    bytes.extend(gif_delay(20));
    bytes.extend(gif_frame(true));
    bytes.push(0x3B);

    let timeline = FrameTimeline::new(gif_delays(&bytes).expect("a walkable GIF"));
    assert_eq!(timeline.frames(), 2, "two descriptors are two frames");
    assert!(
        (timeline.seconds() - 0.3).abs() < 0.001,
        "ten hundredths and twenty of them, and not the frames' own count: {}",
        timeline.seconds()
    );
}

#[test]
fn a_frame_the_file_holds_long_is_the_one_a_late_share_lands_on() {
    let timeline = timeline_of(&[100, 100, 2000]);

    assert_eq!(timeline.frames(), 3);
    assert_eq!(
        timeline.index_at_share(0.0),
        0,
        "the start of the bar is the first frame"
    );
    assert_eq!(
        timeline.index_at_share(1.0),
        2,
        "and the end of it is the last"
    );
    // A share is read as a second of the file's own time: half way through a file whose
    // last frame holds two of its two and a fifth seconds is inside that last frame, and
    // not the middle frame a count of frames would name.
    assert_eq!(timeline.index_at_seconds(0.05), 0);
    assert_eq!(timeline.index_at_seconds(0.15), 1);
    assert_eq!(timeline.index_at_seconds(1.0), 2);
    assert_eq!(
        timeline.index_at_share(0.5),
        2,
        "the frame holding the back half of the file is the one a hand halfway along \
         the bar lands on"
    );
}

#[test]
fn a_share_past_the_end_is_the_last_frame_rather_a_frame_that_is_not_there() {
    let timeline = timeline_of(&[50, 50]);

    assert_eq!(timeline.index_at_share(1.4), 1);
    assert_eq!(timeline.index_at_share(-0.2), 0);
    assert_eq!(timeline.index_at_seconds(9.0), 1);
}

#[test]
fn the_rate_is_the_files_own_pace() {
    let timeline = timeline_of(&[100, 100, 100, 100, 100, 100, 100, 100, 100, 100]);

    assert!(
        (timeline.rate() - 10.0).abs() < 0.001,
        "ten frames of a tenth of a second each is ten a second: {}",
        timeline.rate()
    );
}

#[test]
fn an_animated_webp_reports_the_durations_its_frame_chunks_carry() {
    let bytes = webp_with(&[
        (b"VP8X", vec![0u8; 10]),
        (b"ANIM", vec![0u8; 6]),
        (b"ANMF", webp_frame(40)),
        (b"ANMF", webp_frame(60)),
    ]);

    assert!(
        (timeline_of(&webp_delays(&bytes).expect("a walkable WebP")).seconds() - 0.1).abs() < 0.001,
        "forty and sixty milliseconds is a tenth of a second"
    );
    assert_eq!(webp_delays(&bytes).expect("walked").len(), 2);
}

#[test]
fn an_apng_reports_the_delays_its_frame_controls_carry() {
    let mut bytes = apng_head(2);
    bytes.extend(apng_frame_control(1, 10));
    bytes.extend(apng_frame_control(3, 100));
    bytes.extend_from_slice(&[0u8; 8]);
    bytes.extend_from_slice(b"IEND");
    bytes.extend_from_slice(&[0u8; 4]);

    let delays = apng_delays(&bytes).expect("a walkable APNG");
    assert_eq!(delays, vec![100, 30], "a tenth and three hundredths");
    assert!(
        (timeline_of(&delays).seconds() - 0.13).abs() < 0.001,
        "and the whole of them is a hundred and thirty milliseconds"
    );
}

#[test]
fn a_container_that_runs_out_mid_block_answers_nothing_rather_half_a_timeline() {
    let mut gif = gif_head(0x81);
    gif.extend(gif_delay(10));
    // A descriptor whose local colour table is declared but whose bytes are not there.
    gif.extend_from_slice(&[0x2C, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x02, 0x00, 0x87]);
    assert_eq!(
        gif_delays(&gif),
        None,
        "a block the file does not hold is not a frame it plays"
    );
}

#[test]
fn a_still_picture_has_no_timeline_at_all() {
    let plain = webp_with(&[(b"VP8 ", vec![0u8; 8])]);

    assert_eq!(
        webp_delays(&plain),
        Some(Vec::new()),
        "no frame chunks, and no frames to drag a bar across"
    );
    assert!(
        timeline_of(&[]).frames() == 0 && timeline_of(&[]).rate() <= 1.0,
        "and a timeline of nothing is not a rate of infinity"
    );
}
