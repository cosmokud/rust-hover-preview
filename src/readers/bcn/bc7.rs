//! BC7: eight modes of endpoints and indices, and the format's own tables of them — which
//! mode a block names, how its sixteen texels are cut into subsets, which texel of a
//! subset carries an index one bit shorter, and how much of an endpoint an index weighs.
//!
//! BC6H interpolates with the same weights and asks the same partitions which subset a
//! texel is in, so `FACTORS` and `subset_of` below are this file's to keep and BC6H's to
//! read.

use super::block_formats::Bits;

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
pub(super) static FACTORS: [[u8; 16]; 3] = [
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
        3, 3, 15, 15, 8, 3, 15, 15, 8, 8, 6, 6, 6, 5, 3, 3, 3, 3, 8, 15, 3, 3, 6, 10, 5, 8, 8, 6,
        8, 5, 15, 15, 8, 15, 3, 5, 6, 10, 8, 15, 15, 3, 15, 5, 15, 15, 15, 15, 3, 15, 5, 5, 5, 8,
        5, 10, 5, 10, 8, 13, 15, 12, 3, 3,
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
pub(super) fn bc7(block: &[u8]) -> [[u8; 4]; 16] {
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
pub(super) fn subset_of(subsets: usize, partition: usize, index: usize) -> (usize, bool) {
    match subsets {
        2 => {
            let subset = ((PARTITIONS_2[partition] >> index) & 1) as usize;
            let anchor = if subset != 0 { ANCHORS_2[partition] } else { 0 };

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
