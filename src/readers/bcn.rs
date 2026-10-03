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
