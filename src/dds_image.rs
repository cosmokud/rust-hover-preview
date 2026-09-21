//! DDS textures: everything of one this app can read for itself.
//!
//! A `.dds` is a container a game, a modding tool or an engine writes a texture into, and
//! what is inside it is one of two things. The classic files hold blocks compressed by
//! S3TC — DXT1, DXT3 and DXT5, which Direct3D 11 calls BC1, BC2 and BC3 — with the alpha
//! the DXT3 and DXT5 variants carry. Everything written since the middle of the last
//! decade holds one of the formats those were joined by: BC4 and BC5 for the single and
//! dual channel data a normal map or a mask is, BC6H for a range wider than the display's,
//! and BC7 for a photograph-quality colour at half of BC1's size. And some tools write a
//! texture with no compression at all, in a format declared by channel masks.
//!
//! Both kinds come signed as well: a `_SNORM` sample is a number either side of zero
//! rather than a level — the x and y of a normal map, a height that goes below its plane,
//! the difference between two renders — and a file of one is a file of data rather than of
//! a picture. What is drawn for it is the remap such data is shown with everywhere: the
//! range stretched over the display's, dark below zero and light above it. The conversion
//! that does it lives in `bcn` beside the signed codebook, which is the same conversion
//! applied to the values a block holds.
//!
//! Light comes signed as well, and is drawn the other way round. A signed `BC6H` holds a
//! range that goes below zero rather than one that starts at it — a render's negative
//! radiance, a displacement that goes under its plane — and what is made of it is the curve
//! every other kind of light is put through rather than a stretch over the display's, which
//! lands a light below zero on black. That is the one place a signed file is not drawn as
//! the data it holds: a level is a level whatever its sign, but light is light, and a light
//! that is not there is what a display has to draw it as (see `tone_map`).
//!
//! What Windows answers for is smaller than that. Its DDS codec — the one a `.dds` is
//! opened by before this module is reached (see `wic_image`) — reads BC1, BC2 and BC3,
//! measured rather than assumed: every DXGI format from BC4 on, and every uncompressed
//! one, in both the legacy and the DX10 header styles, is refused with "the image header
//! is unrecognized". What it does give is a decode at the size of the preview rather than
//! at the size of the file, which is why it is asked first and why the three formats it
//! and this module both read do not come here at all.
//!
//! So this module is the reader for the rest of them: the uncompressed formats, whose
//! pixels are read out of the masks or the DXGI format the file declares, and every block
//! format the codec does not read — BC4, BC5, BC6H and BC7 — which are decoded by this app
//! rather than by anything the machine has. Nothing of theirs is here: a block format is
//! arithmetic on sixteen bytes with nothing to carry between blocks, so all six of them
//! live in `bcn` and this module is the container around them — which format a file holds,
//! where the level that format is read from begins, and what becomes of what comes out.
//!
//! Two kinds of format that hold something other than colour are read here as well. The
//! packed HDR ones spend a word on three partial floats rather than four bytes on each of
//! them, which is what light is written as where a half float a channel is more than the
//! picture needs; and a depth buffer is a picture of the depth of a scene, so it is read as
//! the one channel it is and drawn the way the colour format with the same packing is drawn
//! rather than answered with no preview.
//!
//! One thing a `DX10` header declares that a classic one cannot is what a file's alpha
//! channel *means*, and one of those declarations is acted on here: a file that says it is
//! opaque is drawn as if every texel were. A file that says so is a file whose alpha holds
//! nothing — a texture a tool wrote without ever touching the channel — and drawing that
//! channel as it stands is a picture nothing can be seen through, which is a faithful
//! decode and a useless preview. What is *not* acted on is the rest of that declaration: a
//! straight or a premultiplied channel is composited as it always was, which for a
//! premultiplied file means its colour is drawn at the strength its alpha says rather than
//! the strength the colour already holds.
//!
//! What is read of a file is its first face and one mip level of it, which is the shape the
//! rest of the app takes with a picture that holds more than one: a cubemap is six faces
//! and a texture array is however many slices its header declares, and what a hover shows
//! of either is the first one. Which level is the file's own answer to what the texture
//! looks like at a size other than its own, and what a preview draws is a fraction of that
//! size — so the level read is the smallest one the file holds that is still at least as
//! large as the picture being drawn, which for a four-thousand-texel texture previewed
//! eight hundred wide is its third, the same picture to the eye for a sixteenth of the work
//! (see `preview_level`). Nothing is decompressed that is not drawn, and nothing is read
//! that the picture is not made of: the header is a hundred and forty-eight bytes, and what
//! is read after it is one level's own bytes and nothing further — so a cubemap with ten mip
//! levels costs what one face's one level costs, and a file that says it holds more than it
//! does is answered by a short read rather than by an allocation the size of what it claimed.
//!
//! Nothing here is allowed to take the app down with it. A file that is not a DDS, a
//! format that is not one of the ones above, a level that will not fit the budget every
//! other reader is handed, and a read that comes back short are one answer: no preview.

use crate::bcn::{self, Blocks, Texels};
use crate::config::{decode_budget_bytes, frame_bytes_within_budget};
use crate::tone_map::ToneMap;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

/// The four bytes every DDS begins with, which is `'D' 'D' 'S' ' '` little-endian.
const MAGIC: u32 = 0x2053_4444;
/// The header the magic is followed by.
const HEADER_BYTES: usize = 128;
/// The block a file with a `DX10` four-character code carries after its header.
const DX10_HEADER_BYTES: usize = 20;
/// What a pixel format's flags say about where its channels are.
const DDPF_ALPHA: u32 = 0x2;
const DDPF_FOURCC: u32 = 0x4;
const DDPF_RGB: u32 = 0x40;
const DDPF_LUMINANCE: u32 = 0x20_000;
/// What a `DX10` header says its alpha channel means, in the low three bits of
/// `miscFlags2`. Only one of the modes is acted on: a file that calls itself opaque is one
/// whose alpha holds nothing, and it is drawn as every other opaque picture is.
const DDS_ALPHA_MODE_MASK: u32 = 0x7;
const DDS_ALPHA_MODE_OPAQUE: u32 = 3;

/// Whether a file is a texture, as its name says.
///
/// It is the name and not the bytes, for the reason the gates are: what the question is
/// asked for is not whether a file will decode but what the preview of it is drawn over,
/// and that is a question about the kind of file it is — see the tray's `Background`
/// submenu, which keeps a texture's backdrop apart from a picture's.
pub fn is_dds_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("dds"))
}

/// A texture's own size, which is the size the layout places it at.
///
/// `None` is a file that is not a DDS, one whose format neither this module nor `bcn` has
/// a decoder for — a DXGI format outside the two tables below, a pixel format declared no
/// way a pixel can be read out of — and one whose header will not be read, and every one
/// of them is the same answer: no preview.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let header = read_header(path)?;
    Some((header.width, header.height))
}

/// A texture decoded to exactly `width` by `height`, in the order the preview's own frame
/// is composed in: BGRA, top-down, four bytes to the pixel.
///
/// The scaling is done here rather than by a codec, because the format this module reads
/// is one no codec of the machine will do it for: what comes out of a block is a 4x4
/// square of texels, and the whole of the picture has to be decoded before one pixel of
/// the preview is known. What is decoded is the level nearest this size rather than the
/// file's first, so the work is the size of the preview rather than the size of the
/// texture (see `preview_level`).
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    let header = read_header(path)?;

    // The budget every other reader is handed before it allocates, asked of the frame the
    // file would be decoded into: this app composes in BGRA wherever the file's own
    // samples are four bytes wide, unless they are wider, in which case what is decoded
    // is smaller than what is drawn.
    frame_bytes_within_budget(header.width, header.height, 4)?;

    // What is read is one mip level of the first face, and which level that is belongs to
    // the size the picture is being drawn at rather than to the file (see `preview_level`).
    let (level_width, level_height, offset, level_bytes) =
        level_at(&header, preview_level(&header, width, height, path))?;
    let level = read_level(path, offset, level_bytes)?;

    // Whether a curve is applied belongs to the picture rather than to the format — a
    // texture that holds light is brought into a frame with one, and a texture that holds
    // levels is not — so it is read once here and handed to whichever of the two readers
    // this file turns out to be (see `tone_map`).
    let tone = ToneMap::current();

    let mut pixels = match header.picture {
        Picture::Blocks(blocks) => {
            decode_blocks(blocks, &level, level_width, level_height, tone)?
        }
        Picture::Samples(samples) => {
            decode_samples(samples, &level, level_width, level_height, tone)?
        }
    };

    // A file that calls itself opaque is drawn as every other opaque picture is: what its
    // alpha channel holds is nothing, whatever the texels happen to say (see the module's
    // own note).
    if header.opaque {
        for texel in pixels.chunks_exact_mut(4) {
            texel[3] = 255;
        }
    }

    let source = image::RgbaImage::from_raw(level_width, level_height, pixels)?;
    let scaled = if (level_width, level_height) == (width, height) {
        source
    } else {
        image::imageops::resize(&source, width, height, image::imageops::FilterType::Triangle)
    };

    Some(crate::preview_window::rgba_to_bgra(scaled.as_raw()))
}

/// What a file holds: blocks, or samples.
#[derive(Clone, Copy)]
enum Picture {
    Blocks(Blocks),
    Samples(Samples),
}

/// One sample as the file stores it.
#[derive(Clone, Copy)]
enum Sample {
    U8,
    U16,
    F16,
    F32,
    /// A byte read as the signed number it is rather than as a level: one channel of a
    /// file whose data goes both ways from zero — the difference between two renders, or a
    /// value a shader reads and puts through a curve of its own.
    Snorm8,
    /// The same at sixteen bits, where the value is worth keeping that far apart.
    Snorm16,
}

impl Sample {
    fn bytes(self) -> usize {
        match self {
            Sample::U8 | Sample::Snorm8 => 1,
            Sample::U16 | Sample::F16 | Sample::Snorm16 => 2,
            Sample::F32 => 4,
        }
    }

    /// The level this kind of sample holds nothing at, which is what a channel the format
    /// does not carry is drawn with.
    ///
    /// An unsigned sample measures nothing at zero, and so does a float. A signed one
    /// measures it in the middle of its range — a normal map's missing z is a zero like any
    /// other — so a two-channel signed format is drawn the same way the tools that read it
    /// draw it, over the level that means zero rather than over black.
    fn zero_level(self) -> u8 {
        match self {
            Sample::Snorm8 | Sample::Snorm16 => 128,
            Sample::U8 | Sample::U16 | Sample::F16 | Sample::F32 => 0,
        }
    }
}

/// The order a file stores its three colour channels in.
#[derive(Clone, Copy)]
enum Order {
    Rgba,
    Bgra,
    /// Red, green and blue the other way round, with a fourth byte that is not alpha —
    /// the format a texture with nothing transparent in it is written as, where the byte
    /// that would hold an alpha is left to alignment and is drawn as opaque.
    Bgrx,
}

/// A file whose samples are not compressed.
#[derive(Clone, Copy)]
enum Samples {
    /// The formats the legacy header declares: a pixel is a word whose bits belong to
    /// whichever channels the masks name, in whatever widths the masks happen to be.
    Masked {
        bits: u32,
        masks: [u32; 4],
        kind: MaskedKind,
    },
    /// One channel, drawn as grey.
    Red(Sample),
    /// Two channels, drawn with blue at the level the kind measures nothing at and alpha
    /// full, the way a BC5 block is.
    RedGreen(Sample),
    /// Three channels and an alpha, in the order named.
    Colour(Order, Sample),
    /// Ten bits each of red, green and blue and two of alpha, packed into a word.
    Rgb10a2,
    /// Three partial-precision floats packed into a word — eleven bits of red and green, ten
    /// of blue, each with an exponent of its own — which is what light is written as where a
    /// half float a channel is more than the picture needs.
    Rgb11_11_10,
    /// The same idea with one exponent for all three channels: three nine-bit mantissas
    /// sharing the five bits above them.
    Rgb9e5,
    /// Five bits of red and blue and six of green, in that order, with a leading alpha bit
    /// where the format has one.
    Rgb565 { alpha: bool },
}

/// Which of a legacy header's masks describe the picture.
#[derive(Clone, Copy)]
enum MaskedKind {
    Colour,
    Luminance,
    Alpha,
}

impl Samples {
    fn bytes_per_pixel(self) -> Option<usize> {
        match self {
            Samples::Masked { bits, .. } => match bits {
                8 | 16 | 24 | 32 => Some((bits / 8) as usize),
                _ => None,
            },
            Samples::Red(sample) => Some(sample.bytes()),
            Samples::RedGreen(sample) => sample.bytes().checked_mul(2),
            Samples::Colour(_, sample) => sample.bytes().checked_mul(4),
            Samples::Rgb10a2 => Some(4),
            Samples::Rgb11_11_10 => Some(4),
            Samples::Rgb9e5 => Some(4),
            Samples::Rgb565 { .. } => Some(2),
        }
    }
}

impl Picture {
    /// How many bytes the first mip level of the first face is.
    fn level_bytes(self, width: u32, height: u32) -> Option<usize> {
        match self {
            Picture::Blocks(blocks) => {
                let wide = width.div_ceil(4) as usize;
                let high = height.div_ceil(4) as usize;

                wide.checked_mul(high)?.checked_mul(blocks.block_bytes())
            }
            Picture::Samples(samples) => (width as usize)
                .checked_mul(height as usize)?
                .checked_mul(samples.bytes_per_pixel()?),
        }
    }
}

/// A file's header, as far as a preview reads it.
struct Header {
    width: u32,
    height: u32,
    picture: Picture,
    /// Where the first mip level of the first face begins.
    offset: usize,
    /// How many mip levels the file holds, the first of them included: one for a file with
    /// no chain, and no more than the shape allows for one that says it holds more.
    levels: usize,
    /// Whether the file says its alpha channel holds nothing, which is a thing only a
    /// `DX10` header can say (see the module's own note).
    opaque: bool,
}

/// The header of a file whose first bytes claim to be a DDS.
///
/// `None` is every way a file can fail to be one: no magic, a header that does not
/// declare its own size, a shape with no pixels in it, a format this module has no
/// decoder for, and one whose declared pixel width is not one a pixel can be.
fn read_header(path: &Path) -> Option<Header> {
    let mut file = File::open(path).ok()?;

    let mut bytes = [0u8; HEADER_BYTES];
    file.read_exact(&mut bytes).ok()?;

    if u32::from_le_bytes(bytes[0..4].try_into().ok()?) != MAGIC {
        return None;
    }

    // The header names its own size, and a file that does not say 124 is not one this
    // layout can be read out of.
    if u32::from_le_bytes(bytes[4..8].try_into().ok()?) != 124 {
        return None;
    }

    let height = u32::from_le_bytes(bytes[12..16].try_into().ok()?);
    let width = u32::from_le_bytes(bytes[16..20].try_into().ok()?);
    if width == 0 || height == 0 {
        return None;
    }

    let flags = u32::from_le_bytes(bytes[80..84].try_into().ok()?);
    let fourcc: [u8; 4] = bytes[84..88].try_into().ok()?;
    let bits = u32::from_le_bytes(bytes[88..92].try_into().ok()?);
    let masks = [
        u32::from_le_bytes(bytes[92..96].try_into().ok()?),
        u32::from_le_bytes(bytes[96..100].try_into().ok()?),
        u32::from_le_bytes(bytes[100..104].try_into().ok()?),
        u32::from_le_bytes(bytes[104..108].try_into().ok()?),
    ];

    // How many mip levels the file says it holds, which is what a preview may read past the
    // first of. What a file *can* hold is its shape's own answer — halving reaches one texel
    // and stops there — so a count past that is read as the shape's rather than taken at its
    // word, and a count of nothing or of one is a file with no chain at all.
    let mip_count = u32::from_le_bytes(bytes[28..32].try_into().ok()?);
    let levels = mip_count.clamp(1, 32 - width.max(height).leading_zeros()) as usize;

    // The extended header a `DX10` file carries names its format as a DXGI one rather
    // than as masks, and the data starts after it. It is also the only one of the two that
    // says what the alpha channel is for.
    let (picture, offset, opaque) = if fourcc == *b"DX10" {
        let mut extended = [0u8; DX10_HEADER_BYTES];
        file.read_exact(&mut extended).ok()?;

        let format = u32::from_le_bytes(extended[0..4].try_into().ok()?);
        let alpha_mode =
            u32::from_le_bytes(extended[16..20].try_into().ok()?) & DDS_ALPHA_MODE_MASK;

        (
            dx10_picture(format)?,
            HEADER_BYTES + DX10_HEADER_BYTES,
            alpha_mode == DDS_ALPHA_MODE_OPAQUE,
        )
    } else {
        // A classic header declares no such thing: what its alpha means is not written
        // down anywhere, so nothing is read into it.
        (
            legacy_picture(&fourcc, flags, bits, masks)?,
            HEADER_BYTES,
            false,
        )
    };

    // A format whose pixel width is not one a pixel can be written at, and a shape whose
    // first level would not fit an address space, are answered here rather than by a reader
    // that walked into them — and what a level of the file costs is the same arithmetic
    // wherever it is asked for (see `level_at`).
    picture.level_bytes(width, height)?;

    Some(Header {
        width,
        height,
        picture,
        offset,
        levels,
        opaque,
    })
}

/// Where one mip level of the first face is: its shape, where it begins and how many bytes
/// it is.
///
/// The levels lie one after another from the first, each of them the picture before it
/// halved in both directions, so where one begins is the sum of the sizes of the levels in
/// front of it — arithmetic the header's own format and shape are enough for, and checked
/// rather than assumed: a chain whose sum runs past what an address space holds is answered
/// with no level rather than with a sum that wrapped.
fn level_at(header: &Header, level: usize) -> Option<(u32, u32, usize, usize)> {
    let mut width = header.width;
    let mut height = header.height;
    let mut offset = header.offset;

    for _ in 0..level {
        offset = offset.checked_add(header.picture.level_bytes(width, height)?)?;
        width = (width / 2).max(1);
        height = (height / 2).max(1);
    }

    let bytes = header.picture.level_bytes(width, height)?;

    Some((width, height, offset, bytes))
}

/// The level a picture `width` by `height` is read from, which is the smallest one the file
/// holds that is still at least as large as the picture being drawn.
///
/// A mip chain is the file's own answer to what the texture looks like at a size other than
/// its own, and what a preview draws is a fraction of the texture's size, so the level
/// nearest that size is the one to read: a four-thousand-texel texture previewed eight
/// hundred wide is a sixteenth of the blocks to decode this way, and what is read is still
/// larger than what is drawn, so nothing is lost that the preview could have shown. What
/// makes it safe to skip the levels in front is the file: a level is taken only when the
/// file is long enough to hold it, so a header that declares a chain it does not carry is
/// read from its first level, which is the one every file has.
fn preview_level(header: &Header, width: u32, height: u32, path: &Path) -> usize {
    let length = std::fs::metadata(path).map_or(0, |meta| meta.len());
    let mut chosen = 0;

    for level in 1..header.levels {
        let Some((level_width, level_height, offset, bytes)) = level_at(header, level) else {
            break;
        };

        // A level smaller than the picture being drawn holds detail that cannot be shown,
        // and every level past it is smaller still.
        if level_width < width || level_height < height {
            break;
        }

        if offset as u64 + bytes as u64 > length {
            break;
        }

        chosen = level;
    }

    chosen
}

/// The picture a legacy header describes: a four-character code for a compressed format,
/// or the channel masks for an uncompressed one.
fn legacy_picture(fourcc: &[u8; 4], flags: u32, bits: u32, masks: [u32; 4]) -> Option<Picture> {
    if flags & DDPF_FOURCC != 0 {
        return match fourcc {
            b"DXT1" => Some(Picture::Blocks(Blocks::Bc1)),
            // The two premultiplied variants are the same blocks: what differs is whether
            // the colour was multiplied by the alpha before it was written, which is a
            // question for whoever drew it rather than for the decode.
            b"DXT2" | b"DXT3" => Some(Picture::Blocks(Blocks::Bc2)),
            b"DXT4" | b"DXT5" => Some(Picture::Blocks(Blocks::Bc3)),
            b"ATI1" | b"BC4U" => Some(Picture::Blocks(Blocks::Bc4)),
            b"ATI2" | b"BC5U" => Some(Picture::Blocks(Blocks::Bc5)),
            // The signed channels, which are the two above read as numbers rather than as
            // levels and are named apart from them (see `bcn`).
            b"BC4S" => Some(Picture::Blocks(Blocks::Bc4Snorm)),
            b"BC5S" => Some(Picture::Blocks(Blocks::Bc5Snorm)),
            // A legacy header names its format either way: four characters for the ones
            // that have a name, and a `D3DFMT` number for the ones that do not, in the same
            // field — so what no name matched is asked of the numbers.
            _ => numeric_picture(u32::from_le_bytes(*fourcc)),
        };
    }

    // Uncompressed: what the masks say is where the channels are, and no two writers agree
    // on the flags they set beside them, so the kind is read from the flags and the
    // channels are read from the masks.
    let kind = if flags & DDPF_LUMINANCE != 0 {
        MaskedKind::Luminance
    } else if flags & DDPF_RGB != 0 {
        MaskedKind::Colour
    } else if flags & DDPF_ALPHA != 0 {
        MaskedKind::Alpha
    } else {
        return None;
    };

    let described = match kind {
        MaskedKind::Colour => masks[0] != 0 && masks[1] != 0 && masks[2] != 0,
        MaskedKind::Luminance => masks[0] != 0,
        MaskedKind::Alpha => masks[3] != 0,
    };

    if !described {
        return None;
    }

    Some(Picture::Samples(Samples::Masked { bits, masks, kind }))
}

/// The picture a `D3DFMT` number describes, which is the other half of what a legacy
/// header's format field can hold.
///
/// Direct3D 9's own enumeration had no names for the wide and the float layouts — a
/// sixteen-bit-to-the-channel texture, a half float, a float — so a tool of that era wrote
/// the number where a name would have gone, and the `DX10` header that came after names the
/// same layouts as DXGI formats. What this table is, then, is that correspondence and
/// nothing else: each number is read as the DXGI format it stands for, and the two namings
/// of one layout are answered by one decoder rather than by two.
fn numeric_picture(value: u32) -> Option<Picture> {
    dx10_picture(match value {
        // A16B16G16R16, and the signed Q16W16V16U16 beside it.
        36 => 11,
        110 => 13,
        // The half floats: one channel, two, and four.
        111 => 54,
        112 => 34,
        113 => 10,
        // The floats: one channel, two, and four.
        114 => 41,
        115 => 16,
        116 => 2,
        _ => return None,
    })
}

/// The picture a DX10 header's DXGI format describes.
///
/// The numbers are the ones the format is named by; the ranges are the variants of one
/// format, which are a question of how the samples are to be read rather than of how they
/// are stored — a BC1_UNORM_SRGB holds the same blocks as a BC1_UNORM — so all of them
/// are answered by the decoder for the format they are a variant of.
fn dx10_picture(format: u32) -> Option<Picture> {
    match format {
        70..=72 => Some(Picture::Blocks(Blocks::Bc1)),
        73..=75 => Some(Picture::Blocks(Blocks::Bc2)),
        76..=78 => Some(Picture::Blocks(Blocks::Bc3)),
        // BC4 and BC5 are three formats each: the same blocks read as unsigned data, as
        // signed data, and as neither (which is a question for whoever reads them next, and
        // is read here as it is stored).
        79..=81 => Some(Picture::Blocks(match format {
            81 => Blocks::Bc4Snorm,
            _ => Blocks::Bc4,
        })),
        82..=84 => Some(Picture::Blocks(match format {
            84 => Blocks::Bc5Snorm,
            _ => Blocks::Bc5,
        })),
        // BC6H is two formats — the same blocks, read as a range above zero or as a range
        // either side of it — and both are read: a signed file is drawn as light like the
        // unsigned one, with what is below zero landing on black (see `bcn` and `tone_map`).
        95 => Some(Picture::Blocks(Blocks::Bc6h)),
        96 => Some(Picture::Blocks(Blocks::Bc6hSf16)),
        97..=99 => Some(Picture::Blocks(Blocks::Bc7)),

        2 => Some(Picture::Samples(Samples::Colour(Order::Rgba, Sample::F32))),
        10 => Some(Picture::Samples(Samples::Colour(Order::Rgba, Sample::F16))),
        11 => Some(Picture::Samples(Samples::Colour(Order::Rgba, Sample::U16))),
        13 => Some(Picture::Samples(Samples::Colour(Order::Rgba, Sample::Snorm16))),
        16 => Some(Picture::Samples(Samples::RedGreen(Sample::F32))),
        24 => Some(Picture::Samples(Samples::Rgb10a2)),
        26 => Some(Picture::Samples(Samples::Rgb11_11_10)),
        28 | 29 => Some(Picture::Samples(Samples::Colour(Order::Rgba, Sample::U8))),
        31 => Some(Picture::Samples(Samples::Colour(Order::Rgba, Sample::Snorm8))),
        34 => Some(Picture::Samples(Samples::RedGreen(Sample::F16))),
        35 => Some(Picture::Samples(Samples::RedGreen(Sample::U16))),
        37 => Some(Picture::Samples(Samples::RedGreen(Sample::Snorm16))),
        // The depth buffers, read as the one channel they are rather than answered with no
        // preview: a depth texture is a picture of the depth of a scene, and each depth is
        // read as the colour layout with the same packing — the float one as light, the
        // sixteen-bit one as a level — with the stencil a combined format carries beside it
        // left out of the picture, and the stencil-only views not read at all, an index into
        // a stencil buffer being no picture of anything. The enumeration scatters the depth
        // formats among their colour siblings, so they are read where their numbers fall.
        40 => Some(Picture::Samples(Samples::Red(Sample::F32))),
        41 => Some(Picture::Samples(Samples::Red(Sample::F32))),
        45 | 46 => Some(Picture::Samples(Samples::Masked {
            bits: 32,
            masks: [0x00FF_FFFF, 0, 0, 0],
            kind: MaskedKind::Luminance,
        })),
        49 => Some(Picture::Samples(Samples::RedGreen(Sample::U8))),
        51 => Some(Picture::Samples(Samples::RedGreen(Sample::Snorm8))),
        54 => Some(Picture::Samples(Samples::Red(Sample::F16))),
        55 => Some(Picture::Samples(Samples::Red(Sample::U16))),
        56 => Some(Picture::Samples(Samples::Red(Sample::U16))),
        58 => Some(Picture::Samples(Samples::Red(Sample::Snorm16))),
        61 => Some(Picture::Samples(Samples::Red(Sample::U8))),
        63 => Some(Picture::Samples(Samples::Red(Sample::Snorm8))),
        67 => Some(Picture::Samples(Samples::Rgb9e5)),
        85 => Some(Picture::Samples(Samples::Rgb565 { alpha: false })),
        86 => Some(Picture::Samples(Samples::Rgb565 { alpha: true })),
        87 | 91 => Some(Picture::Samples(Samples::Colour(Order::Bgra, Sample::U8))),
        88 => Some(Picture::Samples(Samples::Colour(Order::Bgrx, Sample::U8))),

        _ => None,
    }
}

/// One mip level of the first face, and nothing after it.
///
/// What is read is the one region the preview is made of rather than the file, which is
/// what keeps a cubemap or a mip chain from being paid for: the size is the header's own
/// answer and it is checked against the budget before a byte of it is asked for. What the
/// file itself is, though, is the one thing a header cannot say, so the level is asked of a
/// file long enough to hold it and not of one that is not: a header that describes a level
/// its file does not carry is answered here rather than by an allocation the size of the
/// claim and a read that comes back short.
fn read_level(path: &Path, offset: usize, bytes: usize) -> Option<Vec<u8>> {
    if offset as u64 + bytes as u64 > decode_budget_bytes() {
        return None;
    }

    let mut file = File::open(path).ok()?;

    if file.metadata().ok()?.len() < offset as u64 + bytes as u64 {
        return None;
    }

    file.seek(SeekFrom::Start(offset as u64)).ok()?;

    let mut level = vec![0u8; bytes];
    file.read_exact(&mut level).ok()?;

    Some(level)
}

/// A block-compressed level decoded into the frame the preview composes in.
///
/// Five of the six formats come out of `bcn` as the levels they hold; BC6H comes out as
/// the light it holds, which is not what a frame is composed in — so what that format's
/// texels get is the same curve an EXR's do, applied a texel at a time on the way past
/// (see `tone_map`). A signed file's light below zero is put through the same curve as the
/// rest of it, and what the curve makes of one is black, which is what a display does with
/// light that is not there.
fn decode_blocks(
    blocks: Blocks,
    level: &[u8],
    width: u32,
    height: u32,
    tone: ToneMap,
) -> Option<Vec<u8>> {
    let block_bytes = blocks.block_bytes();
    let wide = width.div_ceil(4) as usize;
    let high = height.div_ceil(4) as usize;

    if level.len() < wide.checked_mul(high)?.checked_mul(block_bytes)? {
        return None;
    }

    let mut pixels = vec![0u8; width as usize * height as usize * 4];
    let stride = width as usize * 4;

    for block_y in 0..high {
        for block_x in 0..wide {
            let start = (block_y * wide + block_x) * block_bytes;
            let texels = match blocks.decode(&level[start..start + block_bytes]) {
                Texels::Levels(levels) => levels,
                Texels::Light(light) => light.map(|texel| {
                    [
                        tone.encode(texel[0]),
                        tone.encode(texel[1]),
                        tone.encode(texel[2]),
                        255,
                    ]
                }),
            };

            for y in 0..4 {
                let row = block_y * 4 + y;
                if row >= height as usize {
                    break;
                }

                for x in 0..4 {
                    let column = block_x * 4 + x;
                    if column >= width as usize {
                        break;
                    }

                    let at = row * stride + column * 4;
                    pixels[at..at + 4].copy_from_slice(&texels[y * 4 + x]);
                }
            }
        }
    }

    Some(pixels)
}

/// An uncompressed level read into the frame the preview composes in.
fn decode_samples(
    samples: Samples,
    level: &[u8],
    width: u32,
    height: u32,
    tone: ToneMap,
) -> Option<Vec<u8>> {
    let count = width as usize * height as usize;
    let stride = samples.bytes_per_pixel()?;

    if level.len() < count.checked_mul(stride)? {
        return None;
    }

    let mut pixels = vec![0u8; count * 4];

    for index in 0..count {
        let pixel = &level[index * stride..][..stride];

        let texel = match samples {
            Samples::Masked { masks, kind, .. } => masked_texel(pixel, masks, kind)?,
            Samples::Red(sample) => {
                let value = sample_level(sample, pixel, tone);
                [value, value, value, 255]
            }
            Samples::RedGreen(sample) => {
                let sample_bytes = sample.bytes();
                [
                    sample_level(sample, &pixel[..sample_bytes], tone),
                    sample_level(sample, &pixel[sample_bytes..], tone),
                    sample.zero_level(),
                    255,
                ]
            }
            Samples::Colour(order, sample) => colour_texel(order, sample, pixel, tone),
            Samples::Rgb10a2 => {
                let packed = u32::from_le_bytes(pixel.try_into().ok()?);

                [
                    narrow_ten(packed & 0x3FF),
                    narrow_ten((packed >> 10) & 0x3FF),
                    narrow_ten((packed >> 20) & 0x3FF),
                    (((packed >> 30) & 0x3) * 85) as u8,
                ]
            }
            Samples::Rgb11_11_10 => {
                let packed = u32::from_le_bytes(pixel.try_into().ok()?);

                [
                    tone.encode(partial_float(packed & 0x7FF, 6)),
                    tone.encode(partial_float((packed >> 11) & 0x7FF, 6)),
                    tone.encode(partial_float((packed >> 22) & 0x3FF, 5)),
                    255,
                ]
            }
            Samples::Rgb9e5 => {
                let packed = u32::from_le_bytes(pixel.try_into().ok()?);
                // One exponent for the three channels, biased by fifteen and eight bits
                // below the mantissa it scales: the format holds no implied leading one, so
                // a channel is its own mantissa times two to the exponent, and the all-zero
                // word is a zero like any other.
                let shared = 2f32.powi((packed >> 27) as i32 - 24);

                [
                    tone.encode((packed & 0x1FF) as f32 * shared),
                    tone.encode(((packed >> 9) & 0x1FF) as f32 * shared),
                    tone.encode(((packed >> 18) & 0x1FF) as f32 * shared),
                    255,
                ]
            }
            Samples::Rgb565 { alpha } => {
                let packed = u16::from_le_bytes(pixel.try_into().ok()?);

                if alpha {
                    [
                        bcn::widen_5(((packed >> 10) & 0x1F) as u8),
                        bcn::widen_5(((packed >> 5) & 0x1F) as u8),
                        bcn::widen_5((packed & 0x1F) as u8),
                        if packed & 0x8000 != 0 { 255 } else { 0 },
                    ]
                } else {
                    bcn::wide_colour(packed)
                }
            }
        };

        pixels[index * 4..index * 4 + 4].copy_from_slice(&texel);
    }

    Some(pixels)
}

/// One pixel of a format whose channels are declared by masks.
fn masked_texel(pixel: &[u8], masks: [u32; 4], kind: MaskedKind) -> Option<[u8; 4]> {
    let mut word = 0u32;
    for (index, byte) in pixel.iter().enumerate() {
        word |= (*byte as u32) << (8 * index);
    }

    let mut texel = [0u8, 0, 0, 255];

    match kind {
        MaskedKind::Colour => {
            texel[0] = masked_channel(word, masks[0])?;
            texel[1] = masked_channel(word, masks[1])?;
            texel[2] = masked_channel(word, masks[2])?;
        }
        MaskedKind::Luminance => {
            let value = masked_channel(word, masks[0])?;
            texel[0] = value;
            texel[1] = value;
            texel[2] = value;
        }
        MaskedKind::Alpha => {
            texel = [0, 0, 0, masked_channel(word, masks[3])?];
        }
    }

    // An alpha mask is optional whatever the kind: a file that declares one and a file
    // that does not are drawn the same way where there is none to read.
    if !matches!(kind, MaskedKind::Alpha) && masks[3] != 0 {
        texel[3] = masked_channel(word, masks[3])?;
    }

    Some(texel)
}

/// One channel out of a packed word, as a level.
///
/// The mask says where the channel is and how wide it is, so a five-bit channel is scaled
/// by its own width rather than by eight — which is what makes one set of masks able to
/// read a 16-bit 565 texture and a 32-bit one alike.
fn masked_channel(word: u32, mask: u32) -> Option<u8> {
    if mask == 0 {
        return None;
    }

    let shift = mask.trailing_zeros();
    let width = (mask >> shift).count_ones();
    let value = ((word & mask) >> shift) as u64;
    let full = (1u64 << width) - 1;

    Some((((value * 255) + full / 2) / full) as u8)
}

/// The three channels and the alpha of a packed colour, in the order it was written in.
fn colour_texel(
    order: Order,
    sample: Sample,
    pixel: &[u8],
    tone: crate::tone_map::ToneMap,
) -> [u8; 4] {
    let width = sample.bytes();

    let texel = |index: usize| &pixel[index * width..][..width];

    match order {
        Order::Rgba => [
            sample_level(sample, texel(0), tone),
            sample_level(sample, texel(1), tone),
            sample_level(sample, texel(2), tone),
            sample_alpha(sample, texel(3)),
        ],
        Order::Bgra => [
            sample_level(sample, texel(2), tone),
            sample_level(sample, texel(1), tone),
            sample_level(sample, texel(0), tone),
            sample_alpha(sample, texel(3)),
        ],
        Order::Bgrx => [
            sample_level(sample, texel(2), tone),
            sample_level(sample, texel(1), tone),
            sample_level(sample, texel(0), tone),
            255,
        ],
    }
}

/// One sample as a level.
///
/// A UNORM holds a level already and is scaled by its own width; a float holds light, and
/// light is what the tone map is for — the same answer an `.exr` or a Radiance `.hdr`
/// gets, because it is the same kind of number (see `tone_map`).
fn sample_level(sample: Sample, bytes: &[u8], tone: crate::tone_map::ToneMap) -> u8 {
    match sample {
        Sample::U8 => bytes[0],
        Sample::U16 => narrow_sixteen(u16::from_le_bytes([bytes[0], bytes[1]])),
        Sample::F16 => tone.encode(bcn::half_to_float(u16::from_le_bytes([bytes[0], bytes[1]]))),
        Sample::F32 => tone.encode(f32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])),
        // Signed data is not light, so no curve is applied to it: it is drawn at the level
        // its own range puts it at, which is what a tool that opens such a file shows.
        Sample::Snorm8 => bcn::snorm8_to_level(bytes[0]),
        Sample::Snorm16 => bcn::snorm16_to_level(u16::from_le_bytes([bytes[0], bytes[1]])),
    }
}

/// One sample as an alpha, which is coverage rather than light: it is brought into the
/// range and scaled, and no curve is applied to it.
fn sample_alpha(sample: Sample, bytes: &[u8]) -> u8 {
    match sample {
        Sample::U8 => bytes[0],
        Sample::U16 => narrow_sixteen(u16::from_le_bytes([bytes[0], bytes[1]])),
        Sample::F16 => narrow_float(bcn::half_to_float(u16::from_le_bytes([bytes[0], bytes[1]]))),
        Sample::F32 => narrow_float(f32::from_le_bytes([
            bytes[0], bytes[1], bytes[2], bytes[3],
        ])),
        // A signed channel is a number either side of zero rather than coverage, and the
        // formats that hold one have no alpha channel at all — the fourth sample of a
        // signed four-channel format is another number, and a preview has nothing to do
        // with it, so the texel is drawn opaque.
        Sample::Snorm8 | Sample::Snorm16 => 255,
    }
}

/// A 16-bit channel as a level.
fn narrow_sixteen(value: u16) -> u8 {
    ((value as u32 * 255 + 32_767) / 65_535) as u8
}

/// A ten-bit channel as a level.
fn narrow_ten(value: u32) -> u8 {
    ((value * 255 + 511) / 1023) as u8
}

/// A number that may be outside the range as a level: clamped, and no curve.
fn narrow_float(value: f32) -> u8 {
    let clamped = if value.is_finite() {
        value.clamp(0.0, 1.0)
    } else {
        0.0
    };

    (clamped * 255.0 + 0.5) as u8
}

/// One partial-precision float out of a word, as the number it is.
///
/// Eleven bits of a word are a five-bit exponent with a six-bit mantissa above it, and ten
/// are the same exponent with five — the two widths the packed format spends on its three
/// channels, which is what lets it hold a range past what a display draws in four bytes to
/// the texel. There is no sign bit to read: the format is light, and light here is not
/// negative. The two ends of the exponent are a half float's own — nothing at all is a zero,
/// and the top of the range is where the format stops — and a mantissa under the smallest
/// ordinary exponent is read as the small number it is rather than flushed to zero, which is
/// what the format's own reference does with one.
fn partial_float(bits: u32, mantissa_bits: u32) -> f32 {
    let exponent = bits >> mantissa_bits;
    let mantissa = bits & ((1 << mantissa_bits) - 1);

    if exponent == 0 {
        mantissa as f32 * 2f32.powi(-14 - mantissa_bits as i32)
    } else if exponent == 0x1F {
        if mantissa == 0 {
            f32::INFINITY
        } else {
            f32::NAN
        }
    } else {
        (1.0 + mantissa as f32 / (1 << mantissa_bits) as f32) * 2f32.powi(exponent as i32 - 15)
    }
}
