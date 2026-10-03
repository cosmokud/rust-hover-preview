//! BC1 to BC5: what a block of the level formats holds, and how the sixteen texels of one
//! come out of it. Beside them are the two things every format here is read through — the
//! bit stream a block is, and the conversions a file's own samples are drawn at — and the
//! shape of the module: which format a block belongs to and how big it is.

use super::bc6h::bc6h;
use super::bc7::bc7;

/// A block as the bit stream the formats read it as: 128 bits, the first byte's low bit
/// first.
///
/// Every field of every one of these formats is a run of bits from that stream, and the
/// runs are not byte-aligned and not in an order a struct could describe — BC6H's are
/// *interleaved* differently in each of its fourteen modes — so what is read is one
/// `peek` and one `read` at a time, out of the block held whole.
pub(super) struct Bits {
    value: u128,
    position: usize,
}

impl Bits {
    pub(super) fn new(block: &[u8]) -> Option<Self> {
        let bytes: [u8; 16] = block.try_into().ok()?;

        Some(Self {
            value: u128::from_le_bytes(bytes),
            position: 0,
        })
    }

    /// The next `bits` bits, which are consumed.
    pub(super) fn read(&mut self, bits: usize) -> u16 {
        let value = self.peek(0, bits);
        self.position += bits;

        value
    }

    /// `bits` bits `offset` past what has been read, which are not consumed.
    ///
    /// A read past the end of the block is answered with zero rather than with a panic:
    /// the last index of a block can be asked for more bits than the block has left, and
    /// what the formats say about that is the same thing as what they say about a block
    /// with no texels in it.
    pub(super) fn peek(&self, offset: usize, bits: usize) -> u16 {
        if bits == 0 {
            return 0;
        }

        let shift = self.position + offset;
        if shift >= 128 {
            return 0;
        }

        (self.value >> shift) as u16 & ((1u32 << bits.min(16)) - 1) as u16
    }
}

/// The sixteen texels of one block, as the format holds them.
///
/// Five of the six formats hold what a screen is addressed with, and one holds light —
/// which is the difference that decides whether a tone map is put over the result, so it
/// is the decoder that says which it is rather than the caller.
pub enum Texels {
    /// Levels, four bytes to the texel.
    Levels([[u8; 4]; 16]),
    /// Light: red, green and blue as the numbers they are, which is what a range wider
    /// than a display's can show — or, in the signed variant, one that goes below zero as
    /// well — is written as.
    Light([[f32; 3]; 16]),
}

/// A block-compressed format, which is four texels square whatever the file's shape is.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Blocks {
    Bc1,
    Bc2,
    Bc3,
    Bc4,
    /// The same blocks as `Bc4`, read as signed data: a single channel either side of zero,
    /// which is what a height map or a difference between two renders is written as.
    Bc4Snorm,
    Bc5,
    /// The same blocks as `Bc5`, read as signed data — the two channels of a normal map as
    /// the numbers they are rather than as levels.
    Bc5Snorm,
    Bc6h,
    /// The same blocks as `Bc6h`, read as signed light: half-float endpoints whose range
    /// goes both ways from zero, which is what a render's negative radiance or a signed
    /// displacement is written as, and which a file names apart from the unsigned one
    /// (`BC6H_SF16` rather than `BC6H_UF16`).
    Bc6hSf16,
    Bc7,
}

impl Blocks {
    /// How many bytes one block of the format is.
    pub fn block_bytes(self) -> usize {
        match self {
            // The ones that carry a colour block alone, and the ones that carry a single
            // channel: eight bytes. The rest hold an alpha block as well, or are nothing
            // but alpha blocks.
            Blocks::Bc1 | Blocks::Bc4 | Blocks::Bc4Snorm => 8,
            Blocks::Bc2
            | Blocks::Bc3
            | Blocks::Bc5
            | Blocks::Bc5Snorm
            | Blocks::Bc6h
            | Blocks::Bc6hSf16
            | Blocks::Bc7 => 16,
        }
    }

    /// The sixteen texels of one block.
    pub fn decode(self, block: &[u8]) -> Texels {
        match self {
            Blocks::Bc1 | Blocks::Bc2 | Blocks::Bc3 => {
                // The colour block is the whole of a BC1 block and the second half of the
                // other two, whose first half is how their alpha is written.
                let colour = if self == Blocks::Bc1 {
                    &block[0..8]
                } else {
                    &block[8..16]
                };

                let mut texels = colour_texels(colour, self == Blocks::Bc1);

                match self {
                    Blocks::Bc1 => {}
                    Blocks::Bc2 => narrow_alpha(&mut texels, &block[0..8]),
                    _ => gradient_alpha(&mut texels, &block[0..8], AlphaChannel::Alpha),
                }

                Texels::Levels(texels)
            }
            Blocks::Bc4 | Blocks::Bc4Snorm => {
                let mut texels = [[0u8, 0, 0, 255]; 16];

                if self == Blocks::Bc4Snorm {
                    gradient_signed(&mut texels, block, AlphaChannel::Red);
                } else {
                    gradient_alpha(&mut texels, block, AlphaChannel::Red);
                }

                // A single channel held in the alpha block's own form, drawn as grey: the
                // value is the value, whichever of the three channels it is asked for as.
                for texel in &mut texels {
                    texel[1] = texel[0];
                    texel[2] = texel[0];
                }

                Texels::Levels(texels)
            }
            Blocks::Bc5 | Blocks::Bc5Snorm => {
                let signed = self == Blocks::Bc5Snorm;

                // The channel a two-channel format does not carry is drawn at the level its
                // kind measures nothing at: black for the unsigned blocks, the middle of the
                // range for the signed ones, because the z a normal map leaves out is a zero
                // rather than an absence.
                let mut texels = [[0u8, 0, if signed { 128 } else { 0 }, 255]; 16];

                if signed {
                    gradient_signed(&mut texels, &block[0..8], AlphaChannel::Red);
                    gradient_signed(&mut texels, &block[8..16], AlphaChannel::Green);
                } else {
                    gradient_alpha(&mut texels, &block[0..8], AlphaChannel::Red);
                    gradient_alpha(&mut texels, &block[8..16], AlphaChannel::Green);
                }

                Texels::Levels(texels)
            }
            Blocks::Bc6h => Texels::Light(bc6h(block, false)),
            Blocks::Bc6hSf16 => Texels::Light(bc6h(block, true)),
            Blocks::Bc7 => Texels::Levels(bc7(block)),
        }
    }
}

/// Which channel a gradient block's values are read into.
#[derive(Clone, Copy)]
enum AlphaChannel {
    Alpha,
    Red,
    Green,
}

/// One sample as the number it is, which for every float here is a half.
pub fn half_to_float(bits: u16) -> f32 {
    let sign = (bits >> 15) as u32;
    let exponent = ((bits >> 10) & 0x1F) as u32;
    let mantissa = (bits & 0x3FF) as u32;

    let value = match exponent {
        // A subnormal half is an ordinary float: the mantissa alone, scaled by the
        // smallest exponent the format has.
        0 => mantissa as f32 * 2.0f32.powi(-24),
        0x1F => {
            if mantissa == 0 {
                f32::INFINITY
            } else {
                f32::NAN
            }
        }
        _ => (1.0 + mantissa as f32 / 1024.0) * 2.0f32.powi(exponent as i32 - 15),
    };

    if sign == 0 {
        value
    } else {
        -value
    }
}

// ---- signed samples: the SNORM side of the same formats ------------------------------
//
// A signed-normalized sample is not a level: it carries a number either side of zero — the
// x and y of a normal, a height that can go below its plane, the difference between two
// renders — and what a preview does with one is the remap every tool that draws such data
// does, stretching `-1..1` over the display's range so a negative value is dark, zero is
// the middle and a positive one is light. A file that holds one is drawn as a picture of
// that data rather than as the data, the same way a BC4 mask is drawn as grey.

/// A signed eight-bit sample as the level it is drawn at.
///
/// The arithmetic is the one DirectXTex and the GPU's own conversion produce — the same
/// answer a file of this kind is given by the tools it was written for — which is what
/// this was checked against rather than derived: the two halves are not quite the
/// symmetric remap they look like, and matching the conversion is worth more here than a
/// tidier formula that disagrees with it by a level.
pub fn snorm8_to_level(value: u8) -> u8 {
    match value.cmp(&128) {
        std::cmp::Ordering::Less => value + 128,
        std::cmp::Ordering::Equal => 0,
        std::cmp::Ordering::Greater => value - 129,
    }
}

/// A signed sixteen-bit sample as the level it is drawn at.
pub fn snorm16_to_level(value: u16) -> u8 {
    let signed = ((value as i16) as f32 / 32767.0).max(-1.0);

    ((signed * 0.5 + 0.5) * 255.0).round() as u8
}

// ---- BC1, BC2 and BC3: two 565 colours and two bits a texel -------------------------

/// The four colours a BC1, BC2 or BC3 block offers and the two bits per texel that choose
/// between them.
fn colour_texels(block: &[u8], bc1: bool) -> [[u8; 4]; 16] {
    let first = u16::from_le_bytes([block[0], block[1]]);
    let second = u16::from_le_bytes([block[2], block[3]]);

    let start = wide_colour(first);
    let end = wide_colour(second);

    // Two of the four are the block's own endpoints and the other two are between them,
    // and a BC1 block whose first endpoint does not sit above its second spends its
    // fourth on a transparent black instead — that format's one bit of alpha, which is
    // why the order they came in is a thing to read rather than a thing to correct. What
    // the third one is, is not the same between them either: the three-colour form has
    // nothing but the two to go on, so it is their midpoint, while the four-colour form
    // gives it three parts of the first endpoint to one of the second.
    let colours: [[u8; 4]; 4] = if bc1 && first <= second {
        [start, end, between(start, end, 1, 1), [0, 0, 0, 0]]
    } else {
        [
            start,
            end,
            between(start, end, 2, 1),
            between(start, end, 1, 2),
        ]
    };

    let mut texels = [[0u8; 4]; 16];
    for row in 0..4 {
        let packed = block[4 + row];

        for column in 0..4 {
            let index = ((packed >> (2 * column)) & 0x3) as usize;
            texels[row * 4 + column] = colours[index];
        }
    }

    texels
}

/// The alpha of a BC2 block, which is four bits to the texel and nothing else.
fn narrow_alpha(texels: &mut [[u8; 4]; 16], block: &[u8]) {
    for (index, byte) in block.iter().enumerate() {
        let low = byte & 0x0F;
        let high = byte >> 4;

        texels[index * 2][3] = (low << 4) | low;
        texels[index * 2 + 1][3] = (high << 4) | high;
    }
}

/// The values of a BC3, BC4 or BC5 gradient block: two endpoints and the six between
/// them, or the four between them and a hard zero and a hard full.
///
/// Which of the two codebooks a block uses is settled by the order its own endpoints came
/// in, and the second is the one that can reach the ends of the range — the reason a
/// gradient block that wrote its larger endpoint first is the one that can say "no alpha
/// at all" and "full alpha" exactly.
fn gradient_alpha(texels: &mut [[u8; 4]; 16], block: &[u8], channel: AlphaChannel) {
    let first = block[0];
    let second = block[1];

    let mut codes = [0u8; 8];
    codes[0] = first;
    codes[1] = second;

    if first <= second {
        for step in 1..5u32 {
            codes[1 + step as usize] =
                (((5 - step) * first as u32 + step * second as u32) / 5) as u8;
        }
        codes[6] = 0;
        codes[7] = 255;
    } else {
        for step in 1..7u32 {
            codes[1 + step as usize] =
                (((7 - step) * first as u32 + step * second as u32) / 7) as u8;
        }
    }

    let mut packed = 0u64;
    for (index, byte) in block[2..8].iter().enumerate() {
        packed |= (*byte as u64) << (8 * index);
    }

    for texel in 0..16 {
        let value = codes[((packed >> (3 * texel)) & 0x7) as usize];

        match channel {
            AlphaChannel::Alpha => texels[texel][3] = value,
            AlphaChannel::Red => texels[texel][0] = value,
            AlphaChannel::Green => texels[texel][1] = value,
        }
    }
}

/// The values of a signed gradient block — BC4_SNORM and each half of BC5_SNORM — as the
/// levels they are drawn at.
///
/// It is the same shape as the unsigned codebook over signed endpoints: the two values a
/// block wrote are numbers either side of zero, the ones between them are the same
/// fractions of the two, and the codebook that reaches the ends of the range holds the
/// ends of the *signed* range rather than nothing and everything. Two differences are
/// worth naming. The endpoints are compared as the signed numbers they are, so a block
/// whose first endpoint is below its second is the one that gets the long codebook — the
/// same rule as the unsigned one, read the other way. And the fractions are rounded rather
/// than truncated, because that is the arithmetic of the reference this path was checked
/// against; the unsigned codebook beside it truncates, and the two differ by at most a
/// level on the blocks where it shows.
fn gradient_signed(texels: &mut [[u8; 4]; 16], block: &[u8], channel: AlphaChannel) {
    // A seventh and a fifth, in the fixed point the fractions are taken at.
    const SEVENTHS: [i32; 6] = [9363, 18724, 28086, 37450, 46812, 56173];
    const FIFTHS: [i32; 4] = [13107, 26215, 39321, 52429];

    // A signed endpoint of the width the format has, with the one value past the end of
    // that range clamped to the end: a signed channel spans `-1..1`, and `-128` is one step
    // beyond it.
    let mut codes = [0i32; 8];
    codes[0] = (block[0] as i8).max(-127) as i32;
    codes[1] = (block[1] as i8).max(-127) as i32;

    if codes[0] > codes[1] {
        for step in 0..6usize {
            codes[2 + step] =
                (SEVENTHS[5 - step] * codes[0] + SEVENTHS[step] * codes[1] + 32768) >> 16;
        }
    } else {
        for step in 0..4usize {
            codes[2 + step] = (FIFTHS[3 - step] * codes[0] + FIFTHS[step] * codes[1] + 32768) >> 16;
        }
        codes[6] = -127;
        codes[7] = 127;
    }

    let mut packed = 0u64;
    for (index, byte) in block[2..8].iter().enumerate() {
        packed |= (*byte as u64) << (8 * index);
    }

    for texel in 0..16 {
        let value = snorm8_to_level(codes[((packed >> (3 * texel)) & 0x7) as usize] as u8);

        match channel {
            AlphaChannel::Alpha => texels[texel][3] = value,
            AlphaChannel::Red => texels[texel][0] = value,
            AlphaChannel::Green => texels[texel][1] = value,
        }
    }
}

/// A 5-bit channel widened to eight.
pub fn widen_5(value: u8) -> u8 {
    (value << 3) | (value >> 2)
}

/// A 6-bit channel widened to eight.
fn widen_6(value: u8) -> u8 {
    (value << 2) | (value >> 4)
}

/// A 565 colour widened to eight bits a channel, with a full alpha.
pub fn wide_colour(packed: u16) -> [u8; 4] {
    [
        widen_5(((packed >> 11) & 0x1F) as u8),
        widen_6(((packed >> 5) & 0x3F) as u8),
        widen_5((packed & 0x1F) as u8),
        255,
    ]
}

/// A colour between two others, weighted: what a block's two middle colours are made of.
fn between(start: [u8; 4], end: [u8; 4], start_weight: u32, end_weight: u32) -> [u8; 4] {
    let total = start_weight + end_weight;
    let mut mixed = [0u8; 4];

    for channel in 0..4 {
        mixed[channel] = ((start_weight * start[channel] as u32 + end_weight * end[channel] as u32)
            / total) as u8;
    }

    mixed
}
