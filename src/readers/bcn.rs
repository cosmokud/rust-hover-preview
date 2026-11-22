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
//! half-float endpoints and a range past what a screen can draw, either side of zero in the
//! signed variant of it — so its texels come back as the numbers they are and are tone
//! mapped on the way to a frame (see `tone_map`).
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
//! which no preview can show. BC6H's signed variant, which was the one format left unread,
//! is read here too and is held to the same test: an independent decoder, four hundred
//! thousand random blocks, every mode, agreeing with it to the bit. What that settles is
//! the signed arithmetic — endpoints read as the numbers of their width rather than as its
//! levels, the signed codebook, the sign a half float carries — since the tables the blocks
//! and the partitions come out of are the unsigned reading's, and it is the unsigned
//! variant that is what holds those to the published ones.
//!
//! The six formats live in three files below: `block_formats` for BC1 to BC5, beside the
//! bit stream and the sample conversions every format here is read through, and `bc7` and
//! `bc6h` for the two whose bit layouts are tables of their own. What is left in this file
//! is the way in: the paths `dds_image` hands a block to.

mod bc6h;
mod bc7;
mod block_formats;

pub(crate) use block_formats::{
    half_to_float, snorm16_to_level, snorm8_to_level, wide_colour, widen_5, Blocks, Texels,
};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn half_to_float_handles_zero_one_negative_and_specials() {
        assert_eq!(half_to_float(0x0000), 0.0);
        assert_eq!(half_to_float(0x3c00), 1.0);
        assert_eq!(half_to_float(0xbc00), -1.0);
        assert!((half_to_float(0x3800) - 0.5).abs() < 1e-6);
        assert!(half_to_float(0x7c00).is_infinite());
        assert!(half_to_float(0x7e00).is_nan());
    }

    #[test]
    fn snorm_and_widen_conversions_cover_extremes() {
        assert_eq!(snorm8_to_level(128), 0);
        assert_eq!(snorm8_to_level(129), 0);
        assert_eq!(snorm8_to_level(0), 128);
        assert_eq!(snorm8_to_level(127), 255);

        assert_eq!(snorm16_to_level(0x8000), 0);
        assert_eq!(snorm16_to_level(0), 128);
        assert_eq!(snorm16_to_level(0x7fff), 255);

        assert_eq!(widen_5(0), 0);
        assert_eq!(widen_5(31), 255);
        assert_eq!(wide_colour(0), [0, 0, 0, 255]);
        assert_eq!(wide_colour(u16::MAX), [255, 255, 255, 255]);
    }

    #[test]
    fn decode_block_sizes_and_solid_bc1_bc4_bc6h_bc7() {
        assert_eq!(Blocks::Bc1.block_bytes(), 8);
        assert_eq!(Blocks::Bc4.block_bytes(), 8);
        assert_eq!(Blocks::Bc7.block_bytes(), 16);
        assert_eq!(Blocks::Bc6h.block_bytes(), 16);

        // Solid white BC1 block: c0 = 0xffff, c1 = 0x0000, all indices = 0 (picks c0).
        let bc1_white = [0xff, 0xff, 0x00, 0x00, 0, 0, 0, 0];
        let Texels::Levels(rgba) = Blocks::Bc1.decode(&bc1_white) else {
            panic!("expected Levels");
        };
        for texel in rgba {
            assert_eq!(texel, [255, 255, 255, 255]);
        }

        // Solid max BC4 unsigned block: e0 = 255, e1 = 0, all indices = 0 (picks e0).
        let bc4_max = [255, 0, 0, 0, 0, 0, 0, 0];
        let Texels::Levels(rgba) = Blocks::Bc4.decode(&bc4_max) else {
            panic!("expected Levels");
        };
        for texel in rgba {
            assert_eq!(texel, [255, 255, 255, 255]);
        }

        // Zeroed BC6H (unsigned + signed) and BC7 blocks decode cleanly without panicking.
        assert!(matches!(Blocks::Bc6h.decode(&[0u8; 16]), Texels::Light(_)));
        assert!(matches!(
            Blocks::Bc6hSf16.decode(&[0u8; 16]),
            Texels::Light(_)
        ));
        assert!(matches!(Blocks::Bc7.decode(&[0u8; 16]), Texels::Levels(_)));
    }
}
