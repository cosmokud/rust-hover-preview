//! Photoshop documents, read for the whole picture the file keeps at its end.
//!
//! A `.psd` is a layered document, and every layer in it is work this app does not
//! do: the pixels, the blend mode, the mask, the clipping, the groups and the effects
//! are all the application's business. It does not have to do any of it, because the
//! format writes the *merged* picture — the whole document as Photoshop shows it when
//! the file is opened — into the last section of the file. So a preview is read by
//! walking past the layers rather than through them: the colour mode section, the
//! image resources section and the layer and mask section are each read for their own
//! length and stepped over, and what is decoded is the one planar picture left.
//!
//! Three things about that picture are what this module is shaped by.
//!
//! It is written at the document's own size, and a layered document can be enormous —
//! thirty thousand pixels to a side is an ordinary thing for a `.psb` to be — so it is
//! decoded straight into the box the layout planned rather than at the size it is and
//! scaled afterwards. Every channel is averaged into that box on the way past, one
//! row of the document at a time, so what a hover holds is the preview itself and a
//! row of the source rather than a picture nobody could hold.
//!
//! Its alpha is unassociated, which is to say the colour of a pixel whose alpha is
//! zero is whatever the layer stack left there and means nothing. Each channel is
//! averaged on its own here, exactly as the picture path averages a PNG, so a target
//! pixel that is partly transparent carries the average of the colours it was made
//! of — the same reading of straight alpha every other picture in this app gets.
//!
//! And it may not be there at all: a document saved with `Maximize Compatibility` off
//! has no merged picture in it, and the answer for one is no preview rather than the
//! layers underneath, because compositing them is the whole of what was skipped. That
//! is also the answer for a colour space this reader does not bring into screen
//! values — Lab and multichannel — and for a file whose own header does not add up: a
//! file this app has no reader for shows nothing, like any other.

use crate::config::{decode_budget_bytes, frame_bytes_within_budget};
use flate2::read::ZlibDecoder;
use std::fs::File;
use std::io::{BufReader, Read, Seek, SeekFrom};
use std::path::Path;

/// What every Photoshop document opens with, and how long the header behind it is.
const SIGNATURE: &[u8; 4] = b"8BPS";
const HEADER_BYTES: usize = 26;
/// The most channels a document may hold: Photoshop's own ceiling, and the bound the
/// row table of a packed picture is read under.
const MAX_CHANNELS: u16 = 56;
/// The longest side a `.psb` may have. A `.psd`'s own ceiling is 30,000; the larger
/// one is what the smaller format was replaced by, and a shape past it is not a
/// document whichever of the two it claims to be.
const MAX_DIMENSION: u32 = 300_000;

/// Whether `path` is named as one of Photoshop's documents, whatever case the name is
/// written in.
pub fn is_psd_file(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.eq_ignore_ascii_case("psd") || ext.eq_ignore_ascii_case("psb"))
        .unwrap_or(false)
}

/// The size of the document `path` holds, read from its header alone.
///
/// A file this reader would refuse is refused here as well, so what the layout places
/// and what the render draws cannot disagree about whether there is a picture at all.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let header = read_header(path)?;
    colour_of(header.color_mode)?;

    Some((header.width, header.height))
}

/// The merged picture `path` holds, decoded into `width` by `height` and handed back
/// as BGRA, top down, the way every frame in this app is composed.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    // The one canvas this reader asks for: a document's own picture is never held
    // whole, so the budget is asked of the preview rather than of the file.
    frame_bytes_within_budget(width, height, 4)?;

    let mut file = BufReader::new(File::open(path).ok()?);
    let mut header_bytes = [0u8; HEADER_BYTES];
    file.read_exact(&mut header_bytes).ok()?;
    let header = header_from(&header_bytes)?;
    let colour = colour_of(header.color_mode)?;

    // The colour mode section: the palette an indexed document's samples are indices
    // into, and a length to step over for every other mode.
    let section = read_length(&mut file, 4)?;
    let palette = match colour {
        Colour::Indexed => Some(read_palette(&mut file, section)?),
        _ => {
            seek_over(&mut file, section)?;
            None
        }
    };

    // The image resources — the colour profile, the thumbnail, the name of the tool
    // that wrote the file — and then the layers, masks, blend modes and effects, which
    // are what this reader is here to walk past. Both are stepped over by the length
    // they declare rather than parsed: the length a writer declares is what its own
    // reader will walk, and a section that is padded still ends where it says it does.
    let resources = read_length(&mut file, 4)?;
    seek_over(&mut file, resources)?;
    let layers = read_length(&mut file, if header.large { 8 } else { 4 })?;
    seek_over(&mut file, layers)?;

    let compression = compression_of(read_u16(&mut file)?)?;

    let bytes_per_sample = match header.depth {
        1 | 8 => 1,
        16 => 2,
        32 => 4,
        _ => return None,
    };
    // A bitmap document is one bit a pixel packed into whole bytes, so its row is an
    // eighth of a picture's; every other depth is whole samples.
    let row_bytes = if header.depth == 1 {
        header.width.div_ceil(8) as usize
    } else {
        header.width as usize * bytes_per_sample
    };

    // A row is read whole into memory whatever happens, so a document whose own width
    // is past the budget is one this hover may not ask for.
    if row_bytes == 0 || row_bytes as u64 > decode_budget_bytes() {
        return None;
    }

    let shape = Shape {
        width: header.width,
        height: header.height,
        depth: header.depth,
        row_bytes,
    };

    // A packed picture writes the length of every row of every channel ahead of the
    // rows themselves, which is what makes a row findable without unpacking the rows
    // before it. The table is read whole — it is the only part of the picture whose
    // size the header alone decides — and is walked in step with the rows.
    let counts = match compression {
        Compression::Rle => read_counts(&mut file, header.channels, header.height, header.large)?,
        _ => Vec::new(),
    };

    // A zipped picture is one stream holding every channel's rows rather than a table
    // of their lengths, so the stream itself is what finds the next row: what comes out
    // of it is the next row of the channel being read, whichever channel that is, and
    // nothing has to be sought to.
    let stream = match compression {
        Compression::Zip { .. } => {
            let mut body = BufReader::new(File::open(path).ok()?);
            body.seek(SeekFrom::Start(file.stream_position().ok()?)).ok()?;
            Some(ZlibDecoder::new(body))
        }
        _ => None,
    };

    let mut picture = Picture {
        compression,
        counts,
        counts_index: 0,
        file,
        stream,
    };

    // The channels of the colour, and then the transparency past them if the document
    // holds one: the first channel after the colour ones is the document's own alpha,
    // and any further channel — a spot colour, a second mask — is not read at all,
    // since the preview is the document and not the inks it was made with.
    let colour_channels = match colour {
        Colour::Rgb => 3,
        Colour::Cmyk => 4,
        _ => 1,
    };
    let has_alpha = header.channels as usize > colour_channels;
    let wanted = colour_channels + usize::from(has_alpha);

    let reducer = Reducer::new(shape.width, shape.height, width, height);
    let mut planes: Vec<Vec<u8>> = Vec::new();

    for channel in 0..wanted {
        let meaning = meaning_of(colour, channel, colour_channels, palette)?;
        let components = match meaning {
            Meaning::Palette(_) => 3,
            Meaning::Component => 1,
        };

        let reduced = reducer.reduce(&mut picture, &shape, &meaning, components)?;

        planes.extend(reduced);
    }

    compose(colour, has_alpha, &planes, width, height)
}

/// The document's own header, read and checked.
struct Header {
    /// Version 2, the large document format: the same picture with wider lengths
    /// around it.
    large: bool,
    channels: u16,
    width: u32,
    height: u32,
    depth: u16,
    color_mode: u16,
}

fn read_header(path: &Path) -> Option<Header> {
    let mut bytes = [0u8; HEADER_BYTES];
    File::open(path).ok()?.read_exact(&mut bytes).ok()?;

    header_from(&bytes)
}

/// The header those bytes hold, checked.
fn header_from(bytes: &[u8]) -> Option<Header> {
    if bytes.len() < HEADER_BYTES || &bytes[0..4] != SIGNATURE {
        return None;
    }

    let large = match u16::from_be_bytes([bytes[4], bytes[5]]) {
        1 => false,
        2 => true,
        _ => return None,
    };

    let header = Header {
        large,
        channels: u16::from_be_bytes([bytes[12], bytes[13]]),
        height: u32::from_be_bytes([bytes[14], bytes[15], bytes[16], bytes[17]]),
        width: u32::from_be_bytes([bytes[18], bytes[19], bytes[20], bytes[21]]),
        depth: u16::from_be_bytes([bytes[22], bytes[23]]),
        color_mode: u16::from_be_bytes([bytes[24], bytes[25]]),
    };

    let sane = header.channels > 0
        && header.channels <= MAX_CHANNELS
        && header.width > 0
        && header.height > 0
        && header.width <= MAX_DIMENSION
        && header.height <= MAX_DIMENSION;

    sane.then_some(header)
}

/// What a document's channels mean, which decides how many of them are colour and how
/// their samples become screen values.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Colour {
    /// One bit a pixel, one channel, white where the bit is set.
    Bitmap,
    /// One channel of grey. A duotone document is read as one too: its own curve is
    /// what makes it a colour picture, and the grey it is built on is what is shown
    /// here rather than nothing at all.
    Gray,
    /// One channel of indices into the palette the colour mode section holds.
    Indexed,
    Rgb,
    Cmyk,
}

fn colour_of(mode: u16) -> Option<Colour> {
    match mode {
        0 => Some(Colour::Bitmap),
        1 | 8 => Some(Colour::Gray),
        2 => Some(Colour::Indexed),
        3 => Some(Colour::Rgb),
        4 => Some(Colour::Cmyk),
        // Multichannel and Lab: the first has no one picture to show, and the second
        // is not brought into screen values here.
        _ => None,
    }
}

/// How the picture at the end of the file is packed.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Compression {
    Raw,
    /// PackBits, one row at a time, with the length of every row written ahead of the
    /// rows in the table this reader holds.
    Rle,
    /// One zlib stream holding every channel's rows in order — the channels one after
    /// another, each in scan-line order — read as the rows themselves rather than
    /// sought into. The predicted form writes a row as the difference between it and
    /// the row before it, which is unwound as it is read.
    Zip { predicted: bool },
}

fn compression_of(value: u16) -> Option<Compression> {
    match value {
        0 => Some(Compression::Raw),
        1 => Some(Compression::Rle),
        2 => Some(Compression::Zip { predicted: false }),
        3 => Some(Compression::Zip { predicted: true }),
        _ => None,
    }
}

/// The shape of one channel's row, and the sample behind it.
struct Shape {
    width: u32,
    height: u32,
    depth: u16,
    row_bytes: usize,
}

/// The rows of the file's picture, read one at a time in the order the file writes
/// them: every row of a channel, then the next channel's.
struct Picture {
    compression: Compression,
    counts: Vec<u32>,
    counts_index: usize,
    file: BufReader<File>,
    /// The one stream a zipped picture is read from, which is the whole of its rows.
    stream: Option<ZlibDecoder<BufReader<File>>>,
}

impl Picture {
    /// The next row of the channel being read, unpacked into `row`.
    fn next_row(&mut self, shape: &Shape, row: &mut Vec<u8>) -> Option<()> {
        match self.compression {
            Compression::Raw => {
                row.clear();
                row.resize(shape.row_bytes, 0);
                self.file.read_exact(row).ok()?;
            }
            Compression::Rle => {
                let count = *self.counts.get(self.counts_index)? as usize;
                self.counts_index += 1;

                // How long a row may be packed to, from the row itself and the runs a
                // PackBits row can grow by. A length past it is not a row, and the
                // bound is what keeps a claim in the table from being an allocation.
                if count > shape.row_bytes + shape.row_bytes / 128 + 2 {
                    return None;
                }

                let mut packed = vec![0u8; count];
                self.file.read_exact(&mut packed).ok()?;
                unpack_row(&packed, row, shape.row_bytes)?;
            }
            Compression::Zip { predicted } => {
                row.clear();
                row.resize(shape.row_bytes, 0);
                self.stream.as_mut()?.read_exact(row).ok()?;

                if predicted {
                    unwind_prediction(row, shape)?;
                }
            }
        }

        Some(())
    }
}

/// Undo the difference a predicted row is stored as.
///
/// A row is written as what each sample is worth more than the sample before it, so
/// what is read back is those added up again — one row of one channel at a time, a
/// channel's rows being what a row is here. An eight-bit sample is a byte and a
/// sixteen-bit one is a big-endian pair, added as the pair it is rather than byte by
/// byte, because what carries between them is part of the sample. A thirty-two-bit row
/// is the one that is not stored in its own order: its samples have their four bytes
/// gathered together before the difference is applied — all of the first, then all of
/// the second — so what is unwound here is the difference first and the gathering
/// after it. A one-bit document is not predicted at all: its rows are packed bits, and
/// a bit cannot be added to.
fn unwind_prediction(row: &mut [u8], shape: &Shape) -> Option<()> {
    match shape.depth {
        8 => {
            for at in 1..row.len() {
                row[at] = row[at].wrapping_add(row[at - 1]);
            }
        }
        16 => {
            for at in (2..row.len()).step_by(2) {
                let previous = u16::from_be_bytes([row[at - 2], row[at - 1]]);
                let sample = u16::from_be_bytes([row[at], row[at + 1]]).wrapping_add(previous);
                row[at..at + 2].copy_from_slice(&sample.to_be_bytes());
            }
        }
        32 => {
            for at in 1..row.len() {
                row[at] = row[at].wrapping_add(row[at - 1]);
            }

            gather_samples(row, shape.width as usize)?;
        }
        _ => return None,
    }

    Some(())
}

/// Put a thirty-two-bit row back into its samples: the four bytes of every sample are
/// written together, of all of them in turn, which is the order the difference above
/// is applied in.
fn gather_samples(row: &mut [u8], width: usize) -> Option<()> {
    if width == 0 || row.len() != width * 4 {
        return None;
    }

    let gathered = row.to_vec();
    for sample in 0..width {
        for byte in 0..4 {
            row[sample * 4 + byte] = gathered[byte * width + sample];
        }
    }

    Some(())
}

/// One row of PackBits, which is how a packed picture writes it: a count that says
/// what the bytes behind it are — `0..=127` meaning the `count + 1` bytes that follow
/// are the row as it is, `-127..=-1` meaning the byte behind it is written `1 - count`
/// times, and `-128` meaning nothing at all — until the row is the length it is.
fn unpack_row(packed: &[u8], row: &mut Vec<u8>, row_bytes: usize) -> Option<()> {
    row.clear();
    row.reserve(row_bytes);

    let mut at = 0;
    while row.len() < row_bytes {
        let count = *packed.get(at)? as i8;
        at += 1;

        if count >= 0 {
            let end = at.checked_add(count as usize + 1)?;
            row.extend_from_slice(packed.get(at..end)?);
            at = end;
        } else if count != -128 {
            let value = *packed.get(at)?;
            at += 1;
            row.resize(row.len() + (1 - count as i32) as usize, value);
        }
    }

    // A run that took the row past its own length is a row that does not add up, and
    // the frame that would have been made of it is not drawn.
    (row.len() == row_bytes).then_some(())
}

/// One channel's row as one value per pixel, which is what every depth is reduced to
/// before it is averaged: eight bits a sample, and what a wider sample holds beyond
/// those bits dropped.
fn row_samples(row: &[u8], shape: &Shape, samples: &mut Vec<u8>) -> Option<()> {
    let width = shape.width as usize;
    samples.clear();
    samples.reserve(width);

    match shape.depth {
        1 => {
            for index in 0..width {
                samples.push(if row.get(index / 8)? & (0x80 >> (index % 8)) != 0 {
                    255
                } else {
                    0
                });
            }
        }
        8 => {
            if row.len() < width {
                return None;
            }
            samples.extend_from_slice(&row[..width]);
        }
        16 => {
            if row.len() < width * 2 {
                return None;
            }
            for index in 0..width {
                samples.push(row[index * 2]);
            }
        }
        32 => {
            if row.len() < width * 4 {
                return None;
            }
            for index in 0..width {
                let at = index * 4;
                let value = f32::from_be_bytes([row[at], row[at + 1], row[at + 2], row[at + 3]]);
                samples.push((value.clamp(0.0, 1.0) * 255.0).round() as u8);
            }
        }
        _ => return None,
    }

    Some(())
}

/// What one channel's samples are: one component of the document's colour, or the
/// palette an index into the picture is looked up in.
///
/// Which component a colour channel is does not have to be said here, because the
/// channels are read in the order the file writes them and the plane a channel leaves
/// behind is the plane its colour is read back from: the colour's own channels first,
/// and the transparency in the one past them, in every mode this reader reads; see
/// `compose`.
#[derive(Clone)]
enum Meaning {
    Component,
    Palette(Box<[u8; 768]>),
}

/// What the channel of a document of `colour_channels` colour channels carries, the
/// first channel past its colour ones being its transparency.
fn meaning_of(
    colour: Colour,
    channel: usize,
    colour_channels: usize,
    palette: Option<[u8; 768]>,
) -> Option<Meaning> {
    if channel == colour_channels {
        return Some(Meaning::Component);
    }

    match colour {
        // An indexed document's one channel is an index into the palette its colour
        // mode section holds, so a document whose palette would not be read is one
        // whose samples cannot be turned into a colour at all.
        Colour::Indexed => palette.map(|palette| Meaning::Palette(Box::new(palette))),
        Colour::Rgb | Colour::Cmyk | Colour::Gray | Colour::Bitmap => Some(Meaning::Component),
    }
}

/// The document's rows averaged into the box the preview is shown in.
///
/// The source is divided into the target rather than sampled from it: every source
/// row and every source column belongs to exactly one target row and one target
/// column, and a target pixel is the average of the block that maps to it. That is
/// what makes a preview the whole of a document at a smaller size rather than a
/// scattering of its pixels, and it is the same box filter the picture path's own
/// resample is.
struct Reducer {
    target_width: usize,
    target_height: usize,
    /// The source columns each target column is the average of, as a start and a
    /// length.
    columns: Vec<(usize, usize)>,
    /// The source rows each target row is the average of, the same way.
    rows: Vec<(usize, usize)>,
}

impl Reducer {
    fn new(source_width: u32, source_height: u32, target_width: u32, target_height: u32) -> Self {
        Self {
            target_width: target_width as usize,
            target_height: target_height as usize,
            columns: spans(source_width, target_width),
            rows: spans(source_height, target_height),
        }
    }

    /// Reduce one channel into `components` planes of eight bits each.
    ///
    /// What is held while it runs is the planes themselves — the size of the preview
    /// rather than the size of the document — and one row group of sums beside them,
    /// so a channel of any size costs the same as the picture it is reduced into.
    ///
    /// The source is walked once whatever the target is, and that is what makes a
    /// document smaller than its box work. An enlarged preview has more target rows than
    /// the source has rows, so most of its target rows cover no source row of their own:
    /// what one of those is made of is the source row already in hand, which is the same
    /// rows the columns of an enlargement are averaged from.
    fn reduce(
        &self,
        picture: &mut Picture,
        shape: &Shape,
        meaning: &Meaning,
        components: usize,
    ) -> Option<Vec<Vec<u8>>> {
        let mut planes = vec![vec![0u8; self.target_width * self.target_height]; components];
        let mut sums = vec![0u64; self.target_width * components];
        let mut row = Vec::new();
        let mut samples = Vec::new();
        // The last source row read, kept for the target rows that have none of their own.
        let mut held = Vec::new();
        let mut consumed = 0usize;

        for (output, (start, rows)) in self.rows.iter().enumerate() {
            sums.fill(0);

            let end = start + rows;
            let mut taken = 0;
            while consumed < end {
                picture.next_row(shape, &mut row)?;
                row_samples(&row, shape, &mut samples)?;
                accumulate(&samples, &self.columns, meaning, &mut sums);
                std::mem::swap(&mut samples, &mut held);
                consumed += 1;
                taken += 1;
            }

            if taken == 0 {
                accumulate(&held, &self.columns, meaning, &mut sums);
                taken = 1;
            }

            let across = taken as u64;
            for column in 0..self.target_width {
                let divisor = across * self.columns[column].1 as u64;
                if divisor == 0 {
                    continue;
                }

                for component in 0..components {
                    let sum = sums[column * components + component];
                    planes[component][output * self.target_width + column] =
                        (sum / divisor).min(255) as u8;
                }
            }
        }

        Some(planes)
    }
}

/// The source indices each target index is the average of: a start and a length, with
/// every source index in exactly one target's span.
fn spans(source: u32, target: u32) -> Vec<(usize, usize)> {
    let source = source.max(1) as usize;
    let target = target.max(1) as usize;

    (0..target)
        .map(|index| {
            let start = (index * source / target).min(source - 1);
            let end = (((index + 1) * source) / target)
                .max(start + 1)
                .min(source);
            (start, end - start)
        })
        .collect()
}

/// Add one source row's samples into the sums of the target columns they belong to.
fn accumulate(
    samples: &[u8],
    columns: &[(usize, usize)],
    meaning: &Meaning,
    sums: &mut [u64],
) {
    match meaning {
        Meaning::Component => {
            for (column, (start, length)) in columns.iter().enumerate() {
                let total: u64 = samples[*start..*start + *length]
                    .iter()
                    .map(|sample| *sample as u64)
                    .sum();
                sums[column] += total;
            }
        }
        // An index is not a colour, so an indexed document is averaged as the colours
        // its palette holds rather than as the indices themselves: the average of two
        // indices is a third entry of the palette, and says nothing about the picture.
        Meaning::Palette(palette) => {
            for (column, (start, length)) in columns.iter().enumerate() {
                for sample in &samples[*start..*start + *length] {
                    let at = *sample as usize * 3;
                    sums[column * 3] += palette[at] as u64;
                    sums[column * 3 + 1] += palette[at + 1] as u64;
                    sums[column * 3 + 2] += palette[at + 2] as u64;
                }
            }
        }
    }
}

/// The reduced planes as the frame this app composes in: BGRA, top down, with the
/// colour's own channels brought into screen values here.
fn compose(
    colour: Colour,
    has_alpha: bool,
    planes: &[Vec<u8>],
    width: u32,
    height: u32,
) -> Option<Vec<u8>> {
    let count = width as usize * height as usize;
    let mut frame = Vec::with_capacity(count * 4);

    for index in 0..count {
        let (red, green, blue) = match colour {
            Colour::Rgb | Colour::Indexed => (
                *planes.first()?.get(index)?,
                *planes.get(1)?.get(index)?,
                *planes.get(2)?.get(index)?,
            ),
            Colour::Gray | Colour::Bitmap => {
                let grey = *planes.first()?.get(index)?;
                (grey, grey, grey)
            }
            // A separated picture is four inks rather than three colours, and the inks
            // it is written as are the light each of them takes away. What is left of
            // the light is what is drawn, which is what a print's own proof shows.
            Colour::Cmyk => {
                let black = 255 - *planes.get(3)?.get(index)? as u32;
                (
                    ((255 - *planes.first()?.get(index)? as u32) * black / 255) as u8,
                    ((255 - *planes.get(1)?.get(index)? as u32) * black / 255) as u8,
                    ((255 - *planes.get(2)?.get(index)? as u32) * black / 255) as u8,
                )
            }
        };

        let alpha = if has_alpha {
            *planes.last()?.get(index)?
        } else {
            255
        };

        frame.extend_from_slice(&[blue, green, red, alpha]);
    }

    Some(frame)
}

/// The palette an indexed document's samples are indices into, read from the colour
/// mode section — which is what that section holds for this mode, and is stepped over
/// for every other.
fn read_palette(file: &mut BufReader<File>, length: u64) -> Option<[u8; 768]> {
    let mut palette = [0u8; 768];
    let take = length.min(768) as usize;
    file.read_exact(&mut palette[..take]).ok()?;
    seek_over(file, length - take as u64)?;

    Some(palette)
}

/// The packed length of every row of every channel, which a packed picture keeps ahead
/// of its rows.
fn read_counts(
    file: &mut BufReader<File>,
    channels: u16,
    height: u32,
    large: bool,
) -> Option<Vec<u32>> {
    let entry = if large { 4 } else { 2 };
    let total = channels as u64 * height as u64 * entry;

    if total > decode_budget_bytes() {
        return None;
    }

    let mut bytes = vec![0u8; total as usize];
    file.read_exact(&mut bytes).ok()?;

    Some(match entry {
        2 => bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]) as u32)
            .collect(),
        _ => bytes
            .chunks_exact(4)
            .map(|chunk| u32::from_be_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]))
            .collect(),
    })
}

/// One of the file's length-prefixed sections, in the width a document of this
/// version writes its prefixes in.
fn read_length(file: &mut BufReader<File>, bytes: usize) -> Option<u64> {
    match bytes {
        4 => Some(read_u32(file)? as u64),
        8 => {
            let mut buffer = [0u8; 8];
            file.read_exact(&mut buffer).ok()?;
            Some(u64::from_be_bytes(buffer))
        }
        _ => None,
    }
}

fn read_u16(file: &mut BufReader<File>) -> Option<u16> {
    let mut buffer = [0u8; 2];
    file.read_exact(&mut buffer).ok()?;

    Some(u16::from_be_bytes(buffer))
}

fn read_u32(file: &mut BufReader<File>) -> Option<u32> {
    let mut buffer = [0u8; 4];
    file.read_exact(&mut buffer).ok()?;

    Some(u32::from_be_bytes(buffer))
}

/// Step over a section by its own length: a length past what a file could hold is a
/// length that belongs to nothing, and is refused rather than sought to.
fn seek_over(file: &mut BufReader<File>, length: u64) -> Option<()> {
    if length > i64::MAX as u64 {
        return None;
    }

    file.seek(SeekFrom::Current(length as i64)).ok()?;

    Some(())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A document of the simplest shape there is — three eight-bit channels, written raw,
    /// with no colour mode section, no image resources and no layers — so that what is
    /// left in the file is the picture and nothing else.
    fn write_document(path: &Path, width: u32, height: u32, planes: &[&[u8]]) {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(SIGNATURE);
        bytes.extend_from_slice(&1u16.to_be_bytes());
        bytes.extend_from_slice(&[0u8; 6]);
        bytes.extend_from_slice(&(planes.len() as u16).to_be_bytes());
        bytes.extend_from_slice(&height.to_be_bytes());
        bytes.extend_from_slice(&width.to_be_bytes());
        bytes.extend_from_slice(&8u16.to_be_bytes());
        bytes.extend_from_slice(&3u16.to_be_bytes());
        bytes.extend_from_slice(&0u32.to_be_bytes());
        bytes.extend_from_slice(&0u32.to_be_bytes());
        bytes.extend_from_slice(&0u32.to_be_bytes());
        bytes.extend_from_slice(&0u16.to_be_bytes());
        for plane in planes {
            bytes.extend_from_slice(plane);
        }

        std::fs::write(path, &bytes).unwrap();
    }

    fn temp_path(name: &str) -> std::path::PathBuf {
        std::env::temp_dir().join(name)
    }

    /// A document smaller than the box it is drawn in is enlarged rather than refused:
    /// what a target row covers where there are more target rows than source rows is the
    /// source row it falls in, so the source is still read exactly once.
    #[test]
    fn enlarges_a_document_into_a_box_bigger_than_it() {
        let path = temp_path("rust-hover-preview-psd-smaller.psd");
        write_document(&path, 2, 2, &[&[10, 20, 30, 40], &[0; 4], &[0; 4]]);

        let frame = decode(&path, 4, 4).expect("an enlarged document is a preview");
        std::fs::remove_file(&path).ok();

        assert_eq!(frame.len(), 4 * 4 * 4);
        for y in 0..4u32 {
            for x in 0..4u32 {
                let at = ((y * 4 + x) * 4) as usize;
                let expected = [10u8, 20, 30, 40][((y * 2 / 4) * 2 + (x * 2 / 4)) as usize];
                assert_eq!(
                    frame[at + 2],
                    expected,
                    "the pixel at {x},{y} is the source pixel it falls in"
                );
                assert_eq!(frame[at + 3], 255, "and the document has no transparency");
            }
        }
    }

    /// A document larger than its box is averaged into it as it always was: every source
    /// row belongs to one target row, so the walk of the source is the same one.
    #[test]
    fn averages_a_document_into_a_box_smaller_than_it() {
        let path = temp_path("rust-hover-preview-psd-larger.psd");
        let row: Vec<u8> = (0..16).collect();
        write_document(&path, 4, 4, &[&row, &[0; 16], &[0; 16]]);

        let frame = decode(&path, 2, 2).expect("a reduced document is a preview");
        std::fs::remove_file(&path).ok();

        let red: Vec<u8> = (0..4usize).map(|at| frame[at * 4 + 2]).collect();
        assert_eq!(
            red,
            vec![10 / 4, 18 / 4, 42 / 4, 50 / 4],
            "each target pixel is the average of the block that maps to it"
        );
    }
}
