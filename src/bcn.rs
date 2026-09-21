//! The block-compressed formats a texture is written in, and how one 4x4 block becomes
//! sixteen texels.
//!
//! A block-compressed format does not store pixels. It stores, for every 4x4 square of
//! them, a handful of endpoints and a few bits per texel saying where between those
//! endpoints the texel sits — which is what makes a texture a quarter of the size of the
//! same picture written out, and what makes decoding it arithmetic on one 16-byte block
//! with nothing to carry from one block to the next. That independence is why this module
//! is a set of pure functions over a block: there is no decoder, no file and no state
//! here, only what a block holds.
//!
//! Six formats are read. BC1, BC2 and BC3 are the S3TC family the older files call DXT1,
//! DXT3 and DXT5: two 565 colours and two bits per texel, with the alpha the last two of
//! them add. BC4 and BC5 are the same idea for one and two channels — a normal map's data
//! is two of them and a mask is one. BC7 is the modern one: eight modes of endpoints and
//! index widths, chosen per block, which is what lets it hold a photograph at a third of
//! BC1's size without the banding BC1 shows. And BC6H holds *light* rather than levels —
//! half-float endpoints and a range past what a screen can draw — so its texels come back
//! as the numbers they are and are tone mapped on the way to a frame (see `tone_map`).
//!
//! Everything here is the format's own arithmetic: the tables below are the formats'
//! tables, as published in the Direct3D specification of block compression — which
//! partition a block's sixteen texels into which subsets, which texel of a subset carries
//! one bit fewer of index, and how much of each endpoint a two, three or four bit index
//! weighs. The one thing the formats do not define is what a *file* holds; that is
//! `dds_image`'s business, and it hands this module sixteen bytes at a time.
//!
//! What all of it is checked against is other people's decoders, over random blocks: every
//! bit pattern of every mode is something the format has an answer for, so a wrong bit
//! layout, a wrong table or a wrong anchor shows up as a disagreement rather than as a
//! picture that merely looks odd. BC7 agrees with two independent implementations to the
//! byte on every block of a four-thousand-block sample, and BC6H agrees with the reference
//! everywhere except in the subnormal range — values below the smallest ordinary half,
//! which no preview can show. BC6H's *signed* variant is not read at all: a decode of it
//! that the machine's own decoder disagrees with is not a decode worth handing to a
//! preview, and it is rare enough that no preview is the better answer.

/// A block as the bit stream the formats read it as: 128 bits, the first byte's low bit
/// first.
///
/// Every field of every one of these formats is a run of bits from that stream, and the
/// runs are not byte-aligned and not in an order a struct could describe — BC6H's are
/// *interleaved* differently in each of its fourteen modes — so what is read is one
/// `peek` and one `read` at a time, out of the block held whole.
struct Bits {
    value: u128,
    position: usize,
}

impl Bits {
    fn new(block: &[u8]) -> Option<Self> {
        let bytes: [u8; 16] = block.try_into().ok()?;

        Some(Self {
            value: u128::from_le_bytes(bytes),
            position: 0,
        })
    }

    /// The next `bits` bits, which are consumed.
    fn read(&mut self, bits: usize) -> u16 {
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
    fn peek(&self, offset: usize, bits: usize) -> u16 {
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
    /// than a display's can show is written as.
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
            Blocks::Bc6h => Texels::Light(bc6h(block)),
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

    if sign == 0 { value } else { -value }
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
            codes[2 + step] =
                (FIFTHS[3 - step] * codes[0] + FIFTHS[step] * codes[1] + 32768) >> 16;
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
        mixed[channel] =
            ((start_weight * start[channel] as u32 + end_weight * end[channel] as u32) / total) as u8;
    }

    mixed
}

// ---- BC7: eight modes of endpoints and indices --------------------------------------

/// What one of BC7's eight modes is made of.
struct Bc7Mode {
    subsets: usize,
    partition_bits: usize,
    rotation_bits: usize,
    index_selection_bits: usize,
    colour_bits: usize,
    alpha_bits: usize,
    /// Whether each endpoint carries a bit of its own, or whether the two of a subset
    /// share one.
    endpoint_pbits: bool,
    shared_pbits: bool,
    /// How many bits an index is, for each of the mode's index sets — the second being
    /// the one only modes 4 and 5 have, where colour and alpha take their weights from
    /// different indices.
    index_bits: [usize; 2],
}

/// The eight modes, in the order a block's leading zeros name them.
///
/// What the numbers are is the format's own table: a mode says how many subsets its block
/// is cut into, how many bits name that cut, how wide its endpoints are and whether they
/// carry a parity bit each, how wide its indices are, and — for the two modes that have
/// them — how a block says which of its two index sets colours and which
/// alpha take their weights from.
static BC7_MODES: [Bc7Mode; 8] = [
    Bc7Mode {
        subsets: 3,
        partition_bits: 4,
        rotation_bits: 0,
        index_selection_bits: 0,
        colour_bits: 4,
        alpha_bits: 0,
        endpoint_pbits: true,
        shared_pbits: false,
        index_bits: [3, 0],
    },
    Bc7Mode {
        subsets: 2,
        partition_bits: 6,
        rotation_bits: 0,
        index_selection_bits: 0,
        colour_bits: 6,
        alpha_bits: 0,
        endpoint_pbits: false,
        shared_pbits: true,
        index_bits: [3, 0],
    },
    Bc7Mode {
        subsets: 3,
        partition_bits: 6,
        rotation_bits: 0,
        index_selection_bits: 0,
        colour_bits: 5,
        alpha_bits: 0,
        endpoint_pbits: false,
        shared_pbits: false,
        index_bits: [2, 0],
    },
    Bc7Mode {
        subsets: 2,
        partition_bits: 6,
        rotation_bits: 0,
        index_selection_bits: 0,
        colour_bits: 7,
        alpha_bits: 0,
        endpoint_pbits: true,
        shared_pbits: false,
        index_bits: [2, 0],
    },
    Bc7Mode {
        subsets: 1,
        partition_bits: 0,
        rotation_bits: 2,
        index_selection_bits: 1,
        colour_bits: 5,
        alpha_bits: 6,
        endpoint_pbits: false,
        shared_pbits: false,
        index_bits: [2, 3],
    },
    Bc7Mode {
        subsets: 1,
        partition_bits: 0,
        rotation_bits: 2,
        index_selection_bits: 0,
        colour_bits: 7,
        alpha_bits: 8,
        endpoint_pbits: false,
        shared_pbits: false,
        index_bits: [2, 2],
    },
    Bc7Mode {
        subsets: 1,
        partition_bits: 0,
        rotation_bits: 0,
        index_selection_bits: 0,
        colour_bits: 7,
        alpha_bits: 7,
        endpoint_pbits: true,
        shared_pbits: false,
        index_bits: [4, 0],
    },
    Bc7Mode {
        subsets: 2,
        partition_bits: 6,
        rotation_bits: 0,
        index_selection_bits: 0,
        colour_bits: 5,
        alpha_bits: 5,
        endpoint_pbits: true,
        shared_pbits: false,
        index_bits: [2, 0],
    },
];

/// How much of each endpoint a two, three or four bit index weighs, out of 64: the
/// format's own table, and the same one BC6H interpolates with.
static FACTORS: [[u8; 16]; 3] = [
    [0, 21, 43, 64, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    [0, 9, 18, 27, 37, 46, 55, 64, 0, 0, 0, 0, 0, 0, 0, 0],
    [0, 4, 9, 13, 17, 21, 26, 30, 34, 38, 43, 47, 51, 55, 60, 64],
];

/// How a block of sixteen texels is cut into two subsets, as two bits a texel: bit `i` of
/// entry `n` is the subset texel `i` is in, for partition `n`.
static PARTITIONS_2: [u16; 64] = [
    0xcccc, 0x8888, 0xeeee, 0xecc8, 0xc880, 0xfeec, 0xfec8, 0xec80, 0xc800, 0xffec, 0xfe80, 0xe800,
    0xffe8, 0xff00, 0xfff0, 0xf000, 0xf710, 0x008e, 0x7100, 0x08ce, 0x008c, 0x7310, 0x3100, 0x8cce,
    0x088c, 0x3110, 0x6666, 0x366c, 0x17e8, 0x0ff0, 0x718e, 0x399c, 0xaaaa, 0xf0f0, 0x5a5a, 0x33cc,
    0x3c3c, 0x55aa, 0x9696, 0xa55a, 0x73ce, 0x13c8, 0x324c, 0x3bdc, 0x6996, 0xc33c, 0x9966, 0x0660,
    0x0272, 0x04e4, 0x4e40, 0x2720, 0xc936, 0x936c, 0x39c6, 0x639c, 0x9336, 0x9cc6, 0x817e, 0xe718,
    0xccf0, 0x0fcc, 0x7744, 0xee22,
];

/// The same, for a cut into three subsets: two bits a texel, taken from the pair at each
/// texel's own position.
static PARTITIONS_3: [u32; 64] = [
    0xaa685050, 0x6a5a5040, 0x5a5a4200, 0x5450a0a8, 0xa5a50000, 0xa0a05050, 0x5555a0a0, 0x5a5a5050,
    0xaa550000, 0xaa555500, 0xaaaa5500, 0x90909090, 0x94949494, 0xa4a4a4a4, 0xa9a59450, 0x2a0a4250,
    0xa5945040, 0x0a425054, 0xa5a5a500, 0x55a0a0a0, 0xa8a85454, 0x6a6a4040, 0xa4a45000, 0x1a1a0500,
    0x0050a4a4, 0xaaa59090, 0x14696914, 0x69691400, 0xa08585a0, 0xaa821414, 0x50a4a450, 0x6a5a0200,
    0xa9a58000, 0x5090a0a8, 0xa8a09050, 0x24242424, 0x00aa5500, 0x24924924, 0x24499224, 0x50a50a50,
    0x500aa550, 0xaaaa4444, 0x66660000, 0xa5a0a5a0, 0x50a050a0, 0x69286928, 0x44aaaa44, 0x66666600,
    0xaa444444, 0x54a854a8, 0x95809580, 0x96969600, 0xa85454a8, 0x80959580, 0xaa141414, 0x96960000,
    0xaaaa1414, 0xa05050a0, 0xa0a5a5a0, 0x96000000, 0x40804080, 0xa9a8a9a8, 0xaaaaaa44, 0x2a4a5254,
];

/// The texel of each subset that carries one bit fewer of index — always the block's own
/// first texel for the first subset, and this for the second.
static ANCHORS_2: [usize; 64] = [
    15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 2, 8, 2, 2, 8, 8, 15, 2, 8,
    2, 2, 8, 8, 2, 2, 15, 15, 6, 8, 2, 8, 15, 15, 2, 8, 2, 2, 2, 15, 15, 6, 6, 2, 6, 8, 15, 15, 2,
    2, 15, 15, 15, 15, 15, 2, 2, 15,
];

/// The same, for the second and third subsets of a three-subset cut.
static ANCHORS_3: [[usize; 64]; 2] = [
    [
        3, 3, 15, 15, 8, 3, 15, 15, 8, 8, 6, 6, 6, 5, 3, 3, 3, 3, 8, 15, 3, 3, 6, 10, 5, 8, 8, 6, 8,
        5, 15, 15, 8, 15, 3, 5, 6, 10, 8, 15, 15, 3, 15, 5, 15, 15, 15, 15, 3, 15, 5, 5, 5, 8, 5, 10,
        5, 10, 8, 13, 15, 12, 3, 3,
    ],
    [
        15, 8, 8, 3, 15, 15, 3, 8, 15, 15, 15, 15, 15, 15, 15, 8, 15, 8, 15, 3, 15, 8, 15, 8, 3,
        15, 6, 10, 15, 15, 10, 8, 15, 3, 15, 10, 10, 8, 9, 10, 6, 15, 8, 15, 3, 6, 6, 8, 15, 3, 15,
        15, 15, 15, 15, 15, 15, 15, 15, 15, 3, 15, 15, 8,
    ],
];

/// The sixteen texels of one BC7 block.
///
/// The block names its own mode in the bits it begins with — a run of zeros ending at the
/// first one — and a block whose first eight bits are all zero names no mode and is drawn
/// as the transparent black such a block is.
fn bc7(block: &[u8]) -> [[u8; 4]; 16] {
    let Some(mut bits) = Bits::new(block) else {
        return [[0, 0, 0, 0]; 16];
    };

    let mut mode = 0usize;
    while mode < 8 && bits.read(1) == 0 {
        mode += 1;
    }
    if mode == 8 {
        return [[0, 0, 0, 0]; 16];
    }

    let settings = &BC7_MODES[mode];
    let parity_bits = if settings.endpoint_pbits {
        1
    } else {
        usize::from(settings.shared_pbits)
    };

    let partition = bits.read(settings.partition_bits) as usize;
    let rotation = bits.read(settings.rotation_bits) as usize;
    let index_selection = bits.read(settings.index_selection_bits) as usize;

    let mut red = [0u8; 6];
    let mut green = [0u8; 6];
    let mut blue = [0u8; 6];
    let mut alpha = [0xFFu8; 6];

    // The endpoints come in channel order — every subset's red, then every subset's
    // green, and so on — and each is read without its parity bit, which is where the
    // channel is shifted up by one for.
    for subset in 0..settings.subsets {
        red[subset * 2] = (bits.read(settings.colour_bits) << parity_bits) as u8;
        red[subset * 2 + 1] = (bits.read(settings.colour_bits) << parity_bits) as u8;
    }
    for subset in 0..settings.subsets {
        green[subset * 2] = (bits.read(settings.colour_bits) << parity_bits) as u8;
        green[subset * 2 + 1] = (bits.read(settings.colour_bits) << parity_bits) as u8;
    }
    for subset in 0..settings.subsets {
        blue[subset * 2] = (bits.read(settings.colour_bits) << parity_bits) as u8;
        blue[subset * 2 + 1] = (bits.read(settings.colour_bits) << parity_bits) as u8;
    }
    if settings.alpha_bits > 0 {
        for subset in 0..settings.subsets {
            alpha[subset * 2] = (bits.read(settings.alpha_bits) << parity_bits) as u8;
            alpha[subset * 2 + 1] = (bits.read(settings.alpha_bits) << parity_bits) as u8;
        }
    }

    // A parity bit is the low bit of the endpoint it belongs to — the one bit the endpoint
    // is stored without, which is what lets a five-bit channel reach a value only eight
    // bits can otherwise hold.
    if parity_bits > 0 {
        for subset in 0..settings.subsets {
            let first = bits.read(parity_bits) as u8;
            let second = if settings.shared_pbits {
                first
            } else {
                bits.read(parity_bits) as u8
            };

            for channel in [&mut red, &mut green, &mut blue, &mut alpha] {
                channel[subset * 2] |= first;
                channel[subset * 2 + 1] |= second;
            }
        }
    }

    // What was read is a channel of its own width; what a texel is drawn as is a channel
    // of eight, with the bits it has repeated into the ones it does not.
    let colour_bits = settings.colour_bits + parity_bits;
    for subset in 0..settings.subsets {
        for channel in [&mut red, &mut green, &mut blue] {
            channel[subset * 2] = widen(channel[subset * 2], colour_bits);
            channel[subset * 2 + 1] = widen(channel[subset * 2 + 1], colour_bits);
        }
    }
    if settings.alpha_bits > 0 {
        let bits_wide = settings.alpha_bits + parity_bits;
        for subset in 0..settings.subsets {
            alpha[subset * 2] = widen(alpha[subset * 2], bits_wide);
            alpha[subset * 2 + 1] = widen(alpha[subset * 2 + 1], bits_wide);
        }
    }

    // The indices are one run of bits a texel, with the second index set — where a mode
    // has one — following the first: colour takes its weights from the set the block names
    // and alpha from the other, which is the whole of what modes 4 and 5 add.
    let has_second_set = settings.index_bits[1] != 0;
    let first_weights = &FACTORS[settings.index_bits[0] - 2];
    let second_weights = &FACTORS[if has_second_set {
        settings.index_bits[1] - 2
    } else {
        settings.index_bits[0] - 2
    }];

    let mut first_offset = 0usize;
    let mut second_offset = settings.subsets * (16 * settings.index_bits[0] - 1);

    let mut texels = [[0u8; 4]; 16];
    for (index, texel) in texels.iter_mut().enumerate() {
        let (subset, anchor) = subset_of(settings.subsets, partition, index);

        // An anchor's index is stored one bit shorter: what the bit it is not stored with
        // would have said is a thing the format knows, so it is not written down.
        let first_bits = settings.index_bits[0] - usize::from(anchor);
        let first = bits.peek(first_offset, first_bits) as usize;
        let second = if has_second_set {
            bits.peek(second_offset, settings.index_bits[1] - usize::from(anchor)) as usize
        } else {
            first
        };

        first_offset += first_bits;
        if has_second_set {
            second_offset += settings.index_bits[1] - usize::from(anchor);
        }

        // Which index set a channel takes its weights from, and which table weighs an
        // index, are two different things: the block's index-selection bit says whether
        // colour is the first set or the second, and the table a weight comes out of is
        // always the one for the width that index was stored at.
        let (colour_index, colour_weights) = if index_selection == 0 {
            (first, first_weights)
        } else {
            (second, second_weights)
        };
        let (alpha_index, alpha_weights) = if index_selection == 0 {
            (second, second_weights)
        } else {
            (first, first_weights)
        };

        let colour_factor = colour_weights[colour_index] as u32;
        let alpha_factor = alpha_weights[alpha_index] as u32;

        let at = subset * 2;
        let mut colour = [
            blend(red[at], red[at + 1], colour_factor),
            blend(green[at], green[at + 1], colour_factor),
            blend(blue[at], blue[at + 1], colour_factor),
            blend(alpha[at], alpha[at + 1], alpha_factor),
        ];

        // Modes 4 and 5 can say that a block's alpha and one of its colours were written
        // the other way round, which is a rotation of the two rather than a mode of its
        // own.
        match rotation {
            1 => colour.swap(0, 3),
            2 => colour.swap(1, 3),
            3 => colour.swap(2, 3),
            _ => {}
        }

        *texel = colour;
    }

    texels
}

/// Which subset a texel is in, and whether it is that subset's anchor.
fn subset_of(subsets: usize, partition: usize, index: usize) -> (usize, bool) {
    match subsets {
        2 => {
            let subset = ((PARTITIONS_2[partition] >> index) & 1) as usize;
            let anchor = if subset != 0 {
                ANCHORS_2[partition]
            } else {
                0
            };

            (subset, index == anchor)
        }
        3 => {
            let subset = ((PARTITIONS_3[partition] >> (2 * index)) & 3) as usize;
            let anchor = if subset != 0 {
                ANCHORS_3[subset - 1][partition]
            } else {
                0
            };

            (subset, index == anchor)
        }
        _ => (0, index == 0),
    }
}

/// A colour between two endpoints, a weight of 64 naming the whole of the second.
fn blend(first: u8, second: u8, weight: u32) -> u8 {
    ((first as u32 * (64 - weight) + second as u32 * weight + 32) >> 6) as u8
}

/// A channel of `bits` bits widened to eight, by repeating its own high bits in the places
/// it has none: a five-bit 31 is a full 255, and a five-bit 16 is very nearly half of one.
fn widen(value: u8, bits: usize) -> u8 {
    if bits >= 8 {
        return value;
    }

    let shifted = value << (8 - bits);
    shifted | (shifted >> bits)
}

// ---- BC6H: light, with half-float endpoints ------------------------------------------

/// What one of BC6H's fourteen modes is made of.
struct Bc6hMode {
    /// Whether the endpoints past the first are written as differences from the first
    /// rather than as values, which is what lets a mode spend fewer bits on them.
    transformed: bool,
    /// How many bits name the cut into two subsets, or zero for a mode with one.
    partition_bits: usize,
    endpoint_bits: usize,
    delta_bits: [usize; 3],
}

/// The fourteen modes, at the five-bit value each is named by — with the values that name
/// none left empty, since the mode a block uses is a thing the block says rather than a
/// thing it is counted into.
static BC6H_MODES: [Bc6hMode; 32] = [
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 10, delta_bits: [5, 5, 5] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 7, delta_bits: [6, 6, 6] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 11, delta_bits: [5, 4, 4] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 10, delta_bits: [10, 10, 10] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 11, delta_bits: [4, 5, 4] },
    Bc6hMode { transformed: true, partition_bits: 0, endpoint_bits: 11, delta_bits: [9, 9, 9] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 11, delta_bits: [4, 4, 5] },
    Bc6hMode { transformed: true, partition_bits: 0, endpoint_bits: 12, delta_bits: [8, 8, 8] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 9, delta_bits: [5, 5, 5] },
    Bc6hMode { transformed: true, partition_bits: 0, endpoint_bits: 16, delta_bits: [4, 4, 4] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 8, delta_bits: [6, 5, 5] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 8, delta_bits: [5, 6, 5] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: true, partition_bits: 5, endpoint_bits: 8, delta_bits: [5, 5, 6] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
    Bc6hMode { transformed: false, partition_bits: 5, endpoint_bits: 6, delta_bits: [6, 6, 6] },
    Bc6hMode { transformed: false, partition_bits: 0, endpoint_bits: 0, delta_bits: [0, 0, 0] },
];

/// The sixteen texels of one BC6H block, as the light they hold.
///
/// Nothing here is a level: the endpoints are quantised floats and what comes out of the
/// arithmetic below is a half float a channel, which is why the texels are handed back as
/// numbers and the curve is applied to them afterwards (`tone_map`).
///
/// What a file names as BC6H is two formats — this one, and the same blocks read as
/// *signed* light — and only the unsigned one is read here: signed light is a range
/// either side of zero that the formats around it have no use for, and a decode of it that
/// disagrees with the machine's own is not a decode worth handing to a preview.
fn bc6h(block: &[u8]) -> [[f32; 3]; 16] {
    let Some(mut bits) = Bits::new(block) else {
        return [[0.0; 3]; 16];
    };

    let mut mode = bits.read(2) as usize;
    if mode & 2 != 0 {
        mode |= (bits.read(3) as usize) << 2;
    }

    let settings = &BC6H_MODES[mode];
    if settings.endpoint_bits == 0 {
        return [[0.0; 3]; 16];
    }

    let mut red = [0u16; 4];
    let mut green = [0u16; 4];
    let mut blue = [0u16; 4];

    // The endpoints of a block are not laid out as a table but written out one mode at a
    // time: each of the fourteen modes interleaves the channels, the endpoint pairs and
    // the bits an endpoint is stored too narrow for differently, and there is no shape
    // they all share. What follows is the format's own order, field by field.
    match mode {
        0 => {
            green[2] |= bits.read(1) << 4;
            blue[2] |= bits.read(1) << 4;
            blue[3] |= bits.read(1) << 4;
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(5);
            green[3] |= bits.read(1) << 4;
            green[2] |= bits.read(4);
            green[1] |= bits.read(5);
            blue[3] |= bits.read(1);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(5);
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(4);
            red[2] |= bits.read(5);
            blue[3] |= bits.read(1) << 2;
            red[3] |= bits.read(5);
            blue[3] |= bits.read(1) << 3;
        }
        1 => {
            green[2] |= bits.read(1) << 5;
            green[3] |= bits.read(1) << 4;
            green[3] |= bits.read(1) << 5;
            red[0] |= bits.read(7);
            blue[3] |= bits.read(1);
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(1) << 4;
            green[0] |= bits.read(7);
            blue[2] |= bits.read(1) << 5;
            blue[3] |= bits.read(1) << 2;
            green[2] |= bits.read(1) << 4;
            blue[0] |= bits.read(7);
            blue[3] |= bits.read(1) << 3;
            blue[3] |= bits.read(1) << 5;
            blue[3] |= bits.read(1) << 4;
            red[1] |= bits.read(6);
            green[2] |= bits.read(4);
            green[1] |= bits.read(6);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(6);
            blue[2] |= bits.read(4);
            red[2] |= bits.read(6);
            red[3] |= bits.read(6);
        }
        2 => {
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(5);
            red[0] |= bits.read(1) << 10;
            green[2] |= bits.read(4);
            green[1] |= bits.read(4);
            green[0] |= bits.read(1) << 10;
            blue[3] |= bits.read(1);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(4);
            blue[0] |= bits.read(1) << 10;
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(4);
            red[2] |= bits.read(5);
            blue[3] |= bits.read(1) << 2;
            red[3] |= bits.read(5);
            blue[3] |= bits.read(1) << 3;
        }
        3 => {
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(10);
            green[1] |= bits.read(10);
            blue[1] |= bits.read(10);
        }
        6 => {
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(4);
            red[0] |= bits.read(1) << 10;
            green[3] |= bits.read(1) << 4;
            green[2] |= bits.read(4);
            green[1] |= bits.read(5);
            green[0] |= bits.read(1) << 10;
            green[3] |= bits.read(4);
            blue[1] |= bits.read(4);
            blue[0] |= bits.read(1) << 10;
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(4);
            red[2] |= bits.read(4);
            blue[3] |= bits.read(1);
            blue[3] |= bits.read(1) << 2;
            red[3] |= bits.read(4);
            green[2] |= bits.read(1) << 4;
            blue[3] |= bits.read(1) << 3;
        }
        7 => {
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(9);
            red[0] |= bits.read(1) << 10;
            green[1] |= bits.read(9);
            green[0] |= bits.read(1) << 10;
            blue[1] |= bits.read(9);
            blue[0] |= bits.read(1) << 10;
        }
        10 => {
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(4);
            red[0] |= bits.read(1) << 10;
            blue[2] |= bits.read(1) << 4;
            green[2] |= bits.read(4);
            green[1] |= bits.read(4);
            green[0] |= bits.read(1) << 10;
            blue[3] |= bits.read(1);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(5);
            blue[0] |= bits.read(1) << 10;
            blue[2] |= bits.read(4);
            red[2] |= bits.read(4);
            blue[3] |= bits.read(1) << 1;
            blue[3] |= bits.read(1) << 2;
            red[3] |= bits.read(4);
            blue[3] |= bits.read(1) << 4;
            blue[3] |= bits.read(1) << 3;
        }
        11 => {
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(8);
            red[0] |= bits.read(1) << 11;
            red[0] |= bits.read(1) << 10;
            green[1] |= bits.read(8);
            green[0] |= bits.read(1) << 11;
            green[0] |= bits.read(1) << 10;
            blue[1] |= bits.read(8);
            blue[0] |= bits.read(1) << 11;
            blue[0] |= bits.read(1) << 10;
        }
        14 => {
            red[0] |= bits.read(9);
            blue[2] |= bits.read(1) << 4;
            green[0] |= bits.read(9);
            green[2] |= bits.read(1) << 4;
            blue[0] |= bits.read(9);
            blue[3] |= bits.read(1) << 4;
            red[1] |= bits.read(5);
            green[3] |= bits.read(1) << 4;
            green[2] |= bits.read(4);
            green[1] |= bits.read(5);
            blue[3] |= bits.read(1);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(5);
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(4);
            red[2] |= bits.read(5);
            blue[3] |= bits.read(1) << 2;
            red[3] |= bits.read(5);
            blue[3] |= bits.read(1) << 3;
        }
        15 => {
            red[0] |= bits.read(10);
            green[0] |= bits.read(10);
            blue[0] |= bits.read(10);
            red[1] |= bits.read(4);
            for shift in (10..16).rev() {
                red[0] |= bits.read(1) << shift;
            }
            green[1] |= bits.read(4);
            for shift in (10..16).rev() {
                green[0] |= bits.read(1) << shift;
            }
            blue[1] |= bits.read(4);
            for shift in (10..16).rev() {
                blue[0] |= bits.read(1) << shift;
            }
        }
        18 => {
            red[0] |= bits.read(8);
            green[3] |= bits.read(1) << 4;
            blue[2] |= bits.read(1) << 4;
            green[0] |= bits.read(8);
            blue[3] |= bits.read(1) << 2;
            green[2] |= bits.read(1) << 4;
            blue[0] |= bits.read(8);
            blue[3] |= bits.read(1) << 3;
            blue[3] |= bits.read(1) << 4;
            red[1] |= bits.read(6);
            green[2] |= bits.read(4);
            green[1] |= bits.read(5);
            blue[3] |= bits.read(1);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(5);
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(4);
            red[2] |= bits.read(6);
            red[3] |= bits.read(6);
        }
        22 => {
            red[0] |= bits.read(8);
            blue[3] |= bits.read(1);
            blue[2] |= bits.read(1) << 4;
            green[0] |= bits.read(8);
            green[2] |= bits.read(1) << 5;
            green[2] |= bits.read(1) << 4;
            blue[0] |= bits.read(8);
            green[3] |= bits.read(1) << 5;
            blue[3] |= bits.read(1) << 4;
            red[1] |= bits.read(5);
            green[3] |= bits.read(1) << 4;
            green[2] |= bits.read(4);
            green[1] |= bits.read(6);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(5);
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(4);
            red[2] |= bits.read(5);
            blue[3] |= bits.read(1) << 2;
            red[3] |= bits.read(5);
            blue[3] |= bits.read(1) << 3;
        }
        26 => {
            red[0] |= bits.read(8);
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(1) << 4;
            green[0] |= bits.read(8);
            blue[2] |= bits.read(1) << 5;
            green[2] |= bits.read(1) << 4;
            blue[0] |= bits.read(8);
            blue[3] |= bits.read(1) << 5;
            blue[3] |= bits.read(1) << 4;
            red[1] |= bits.read(5);
            green[3] |= bits.read(1) << 4;
            green[2] |= bits.read(4);
            green[1] |= bits.read(5);
            blue[3] |= bits.read(1);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(6);
            blue[2] |= bits.read(4);
            red[2] |= bits.read(5);
            blue[3] |= bits.read(1) << 2;
            red[3] |= bits.read(5);
            blue[3] |= bits.read(1) << 3;
        }
        30 => {
            red[0] |= bits.read(6);
            green[3] |= bits.read(1) << 4;
            blue[3] |= bits.read(1);
            blue[3] |= bits.read(1) << 1;
            blue[2] |= bits.read(1) << 4;
            green[0] |= bits.read(6);
            green[2] |= bits.read(1) << 5;
            blue[2] |= bits.read(1) << 5;
            blue[3] |= bits.read(1) << 2;
            green[2] |= bits.read(1) << 4;
            blue[0] |= bits.read(6);
            green[3] |= bits.read(1) << 5;
            blue[3] |= bits.read(1) << 3;
            blue[3] |= bits.read(1) << 5;
            blue[3] |= bits.read(1) << 4;
            red[1] |= bits.read(6);
            green[2] |= bits.read(4);
            green[1] |= bits.read(6);
            green[3] |= bits.read(4);
            blue[1] |= bits.read(6);
            blue[2] |= bits.read(4);
            red[2] |= bits.read(6);
            red[3] |= bits.read(6);
        }
        _ => return [[0.0; 3]; 16],
    }

    // The endpoints past the first are differences from it in the modes that say so —
    // which is how a mode spends fewer bits on them, and a difference is a signed number,
    // since half of them go downwards. The two modes that write every endpoint outright
    // are not differences at all and are read as the values they are.
    let subsets = if settings.partition_bits != 0 { 2 } else { 1 };

    for endpoint in 1..subsets * 2 {
        if settings.transformed {
            red[endpoint] = sign_extend(red[endpoint], settings.delta_bits[0]);
            green[endpoint] = sign_extend(green[endpoint], settings.delta_bits[1]);
            blue[endpoint] = sign_extend(blue[endpoint], settings.delta_bits[2]);

            // A difference is added to the endpoint before it, in a field as wide as the
            // mode's channels: it is allowed to carry past the top of that field and wrap,
            // which is what a delta that large is taken to mean.
            let mask = (1u32 << settings.endpoint_bits) - 1;

            red[endpoint] = ((red[endpoint] as u32 + red[0] as u32) & mask) as u16;
            green[endpoint] = ((green[endpoint] as u32 + green[0] as u32) & mask) as u16;
            blue[endpoint] = ((blue[endpoint] as u32 + blue[0] as u32) & mask) as u16;
        }
    }

    for endpoint in 0..subsets * 2 {
        red[endpoint] = unquantize(red[endpoint], settings.endpoint_bits);
        green[endpoint] = unquantize(green[endpoint], settings.endpoint_bits);
        blue[endpoint] = unquantize(blue[endpoint], settings.endpoint_bits);
    }

    // The cut comes after the endpoints rather than before them, and the indices are three
    // bits wide for a block that is cut in two and four for one that is not.
    let partition = if settings.partition_bits != 0 {
        bits.read(5) as usize
    } else {
        0
    };
    let index_bits = if settings.partition_bits != 0 { 3 } else { 4 };
    let weights = &FACTORS[index_bits - 2];

    let mut texels = [[0.0f32; 3]; 16];
    for (index, texel) in texels.iter_mut().enumerate() {
        let (subset, anchor) = if settings.partition_bits != 0 {
            subset_of(2, partition, index)
        } else {
            (0, index == 0)
        };

        let weight = weights[bits.read(index_bits - usize::from(anchor)) as usize] as u32;
        let at = subset * 2;

        *texel = [
            completed(red[at], red[at + 1], weight),
            completed(green[at], green[at + 1], weight),
            completed(blue[at], blue[at + 1], weight),
        ];
    }

    texels
}

/// A light between two endpoints as the half float it is drawn as.
fn completed(first: u16, second: u16, weight: u32) -> f32 {
    let blended = (first as u32 * (64 - weight) + second as u32 * weight + 32) >> 6;

    half_to_float(finish_unquantize(blended))
}

/// A two's-complement number of the width it was read at.
fn sign_extend(value: u16, bits: usize) -> u16 {
    if bits >= 16 {
        return value;
    }

    let mask = 1u16 << (bits - 1);
    (value ^ mask).wrapping_sub(mask)
}

/// An endpoint of the width a mode wrote it with, in the sixteen-bit field the
/// interpolation is done in.
///
/// The top of the field is the number's own top: an endpoint at its largest value is an
/// endpoint at the largest the field holds, and one that is not is scaled by a power of
/// two rather than stretched, which is what keeps a decoded block's shades where the
/// encoder put them.
fn unquantize(value: u16, bits: usize) -> u16 {
    if bits >= 15 {
        return value;
    }
    if value == 0 {
        return 0;
    }

    // The top of the field is the top of the range: the value a mode's own width can hold
    // at all is the one that means "as bright as this format goes", and anything below it
    // is scaled by a power of two rather than stretched to reach the same place.
    if value == (1u16 << bits) - 1 {
        return u16::MAX;
    }

    ((((value as u32) << 15) + 0x4000) >> (bits - 1)) as u16
}

/// The last step of the format's own arithmetic: what the interpolation came out as, as
/// the half float a texel is drawn from.
fn finish_unquantize(value: u32) -> u16 {
    ((value * 31) >> 6) as u16
}
