//! BC6H: light, with half-float endpoints — what one block holds as the numbers a preview
//! tone maps on the way to a frame.
//!
//! The blocks are the same in both readings a file can name, and so are the tables the bit
//! layouts come out of; the three steps the readings differ in are branches of `bc6h`
//! rather than a decoder of their own.

use super::bc7::{subset_of, FACTORS};
use super::block_formats::{half_to_float, Bits};

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
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 10,
        delta_bits: [5, 5, 5],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 7,
        delta_bits: [6, 6, 6],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 11,
        delta_bits: [5, 4, 4],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 10,
        delta_bits: [10, 10, 10],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 11,
        delta_bits: [4, 5, 4],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 0,
        endpoint_bits: 11,
        delta_bits: [9, 9, 9],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 11,
        delta_bits: [4, 4, 5],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 0,
        endpoint_bits: 12,
        delta_bits: [8, 8, 8],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 9,
        delta_bits: [5, 5, 5],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 0,
        endpoint_bits: 16,
        delta_bits: [4, 4, 4],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 8,
        delta_bits: [6, 5, 5],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 8,
        delta_bits: [5, 6, 5],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: true,
        partition_bits: 5,
        endpoint_bits: 8,
        delta_bits: [5, 5, 6],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 5,
        endpoint_bits: 6,
        delta_bits: [6, 6, 6],
    },
    Bc6hMode {
        transformed: false,
        partition_bits: 0,
        endpoint_bits: 0,
        delta_bits: [0, 0, 0],
    },
];

/// The sixteen texels of one BC6H block, as the light they hold.
///
/// Nothing here is a level: the endpoints are quantised floats and what comes out of the
/// arithmetic below is a half float a channel, which is why the texels are handed back as
/// numbers and the curve is applied to them afterwards (`tone_map`).
///
/// What a file names as BC6H is two formats, and both are read here: the blocks are the
/// same, the tables the bit layouts and the partitions come out of are the same, and the
/// three steps the two readings differ in are branches of this one rather than a decoder
/// of their own — a signed file's endpoints are the numbers of the width they were written
/// at rather than the levels of it, its endpoints' arithmetic is the signed codebook beside
/// the unsigned one, and what a light that came out below zero is written as is the sign a
/// half float carries. `signed` says which of the two readings this block is to be given.
pub(super) fn bc6h(block: &[u8], signed: bool) -> [[f32; 3]; 16] {
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

    // A signed file's every endpoint is a number rather than a level, and what the fields
    // were read as is the same either way: the width a mode wrote one at is the width of
    // the number, so what the unsigned reading takes as the value it is, this takes as the
    // number of that width it stands for. The additions above are the format's own in both
    // readings — a difference that carries past the top of the field wraps in it — and what
    // wraps around is the number the wrap landed on, which is why this reads the sum the
    // same way it reads the endpoints that were written outright.
    if signed {
        for endpoint in 0..subsets * 2 {
            red[endpoint] = sign_extend(red[endpoint], settings.endpoint_bits);
            green[endpoint] = sign_extend(green[endpoint], settings.endpoint_bits);
            blue[endpoint] = sign_extend(blue[endpoint], settings.endpoint_bits);
        }
    }

    let mut red_value = [0i32; 4];
    let mut green_value = [0i32; 4];
    let mut blue_value = [0i32; 4];

    for endpoint in 0..subsets * 2 {
        if signed {
            red_value[endpoint] = unquantize_signed(red[endpoint], settings.endpoint_bits);
            green_value[endpoint] = unquantize_signed(green[endpoint], settings.endpoint_bits);
            blue_value[endpoint] = unquantize_signed(blue[endpoint], settings.endpoint_bits);
        } else {
            red_value[endpoint] = unquantize(red[endpoint], settings.endpoint_bits) as i32;
            green_value[endpoint] = unquantize(green[endpoint], settings.endpoint_bits) as i32;
            blue_value[endpoint] = unquantize(blue[endpoint], settings.endpoint_bits) as i32;
        }
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
            completed(red_value[at], red_value[at + 1], weight, signed),
            completed(green_value[at], green_value[at + 1], weight, signed),
            completed(blue_value[at], blue_value[at + 1], weight, signed),
        ];
    }

    texels
}

/// A light between two endpoints as the half float it is drawn as.
///
/// The blend is the format's own and is the same in both readings of it — a weight of 64
/// naming the whole of the second endpoint, the sum taken in a field wide enough for either
/// — since what differs is not how two endpoints are mixed but what they are worth and what
/// the answer is written as, which is `finish_unquantize`'s business.
fn completed(first: i32, second: i32, weight: u32, signed: bool) -> f32 {
    let weight = weight as i32;
    let blended = (first * (64 - weight) + second * weight + 32) >> 6;

    half_to_float(finish_unquantize(blended, signed))
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

/// A signed endpoint of the width a mode wrote it with, as the number either side of zero
/// it stands for.
///
/// It is the unsigned reading of the same field over the range a signed sample has rather
/// than over a level's: the magnitude is scaled by a power of two rather than stretched, and
/// the ends of the range saturate rather than wrap — the most negative a width can write is
/// one step past what a signed channel holds, and it is read as the end of the range the way
/// the largest positive is. Nothing here is a level: what comes out is the number the block
/// stands for, which is what the interpolation and the curve are handed next.
fn unquantize_signed(value: u16, bits: usize) -> i32 {
    if bits >= 16 {
        return value as i16 as i32;
    }

    let number = value as i16 as i32;
    let magnitude = number.unsigned_abs();

    let unquantized = if magnitude == 0 {
        0
    } else if magnitude >= (1 << (bits - 1)) - 1 {
        0x7fff
    } else {
        ((magnitude << 15) + 0x4000) >> (bits - 1)
    };

    if number < 0 {
        -(unquantized as i32)
    } else {
        unquantized as i32
    }
}

/// The last step of the format's own arithmetic: what the interpolation came out as, as
/// the half float a texel is drawn from.
///
/// The two readings of the format part here, over the two fields a half float has. An
/// unsigned light is a magnitude, and it is scaled into the whole of what a half float
/// holds; a signed one is a magnitude and a sign, so the sign is written in the half
/// float's own sign bit and one bit less of the field is left for the magnitude to be
/// scaled into. What is scaled either way is the format's `31/32`, which lands the largest
/// a block can mean on the largest ordinary half. A signed light that came out below zero
/// is kept below zero — that is the number the file holds — and what a preview draws of one
/// is the curve's business rather than this function's (see `tone_map`).
fn finish_unquantize(value: i32, signed: bool) -> u16 {
    if signed {
        let magnitude = (value.unsigned_abs() * 31) >> 5;

        if value < 0 {
            (magnitude as u16) | 0x8000
        } else {
            magnitude as u16
        }
    } else {
        ((value.max(0) as u32 * 31) >> 6) as u16
    }
}
