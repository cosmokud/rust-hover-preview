//! A file as text: reading it, decoding it, and the two formats that are markup
//! rather than text — an RTF's stripped text and ANSI color escapes.
//!
//! The decode is what keeps a renamed archive or executable out of the renderer:
//! bytes that are not text at all are reported as not text, which is a different
//! answer from a file that decodes to nothing.
//!
//! Tabs are expanded here, before anything else looks at the text, so the
//! highlighter, the Markdown parser and the layout all see the same columns.

use super::document::{extension_of, wants_ansi, Span};
use crate::text::text_paint::{blend, readable, rgb, TextStyle};
use crate::text::text_theme::LoadedTheme;
use std::fs::File;
use std::io::Read;
use std::path::Path;

/// Bytes read from the file. A preview shows one screenful, and this bounds the
/// work a hover can trigger on a file that happens to be enormous.
const READ_LIMIT_BYTES: u64 = 2 * 1024 * 1024;
const TAB_WIDTH: usize = 4;
/// A decoded file with its tabs already expanded. Lines are located by offset in
/// `text` rather than copied out of it.
pub(super) struct SourceText {
    pub(super) text: String,
    pub(super) truncated: bool,
}

pub(super) fn read_source(path: &Path) -> Option<SourceText> {
    let file = File::open(path).ok()?;
    let mut bytes = Vec::new();
    let mut limited = file.take(READ_LIMIT_BYTES + 1);
    limited.read_to_end(&mut bytes).ok()?;

    let truncated = bytes.len() as u64 > READ_LIMIT_BYTES;
    bytes.truncate(READ_LIMIT_BYTES as usize);

    let text = decode_text(&bytes, legacy_encoding(&extension_of(path)), truncated)?;
    let text = expand_tabs(&text);

    Some(SourceText { text, truncated })
}

/// The single-byte encoding a file without a byte order mark is assumed to use.
/// Scene NFO art predates code pages being a settled thing, so those files get
/// the table they were drawn in.
#[derive(Clone, Copy)]
enum LegacyEncoding {
    Windows1252,
    Cp437,
}

fn legacy_encoding(extension: &str) -> LegacyEncoding {
    if wants_ansi(extension) {
        LegacyEncoding::Cp437
    } else {
        LegacyEncoding::Windows1252
    }
}

/// Decode a text file: byte order marks first, then UTF-16's own two shapes, then
/// UTF-8, then the legacy code page. `None` means the bytes are not text at all,
/// which is how a renamed archive or executable stays out of the renderer.
fn decode_text(bytes: &[u8], legacy: LegacyEncoding, truncated: bool) -> Option<String> {
    if bytes.is_empty() {
        return Some(String::new());
    }

    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        return Some(String::from_utf8_lossy(&bytes[3..]).into_owned());
    }

    if bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]) {
        return Some(decode_utf16(&bytes[2..], bytes[0] == 0xFE));
    }

    // NUL bytes are checked before the UTF-8 decode, because they are valid
    // UTF-8: a UTF-16 file without a mark and a renamed executable both read as
    // UTF-8 text holding NULs, and neither is a page to show.
    if bytes[..bytes.len().min(8192)].contains(&0) {
        return decode_utf16_without_mark(bytes);
    }

    if let Ok(text) = std::str::from_utf8(bytes) {
        // Valid UTF-8 is not the same as a page: bytes that decode into mostly control characters are
        // a binary file that happened to pass. ANSI-art extensions are exempt, since their own escapes
        // are control characters — and so is a page that is not UTF-8 at all, which is asked below
        // before this answer is taken as a refusal (a UTF-16 page whose low bytes are control
        // characters decodes as UTF-8 holding them, which is not a page either way).
        if matches!(legacy, LegacyEncoding::Cp437) || !is_mostly_control(text) {
            return Some(text.to_string());
        }
    }

    // The byte cap can land inside a multi-byte character. Dropping the few
    // bytes a split character occupies keeps a long UTF-8 file from being
    // mistaken for a legacy-encoded one.
    if truncated {
        let head = &bytes[..bytes.len().saturating_sub(4)];
        if let Ok(text) = std::str::from_utf8(head) {
            return Some(text.to_string());
        }
    }

    // **A page of Chinese, Japanese or Korean with no ASCII in it has no NUL byte either**, so the
    // NUL check above let it through to here. It is not UTF-8 and it is not the code page: it is a
    // UTF-16 file whose every character is above U+00FF.
    if !matches!(legacy, LegacyEncoding::Cp437) {
        if let Some(text) = decode_utf16_without_mark(bytes) {
            return Some(text);
        }
    }

    // The bytes that decoded as mostly control characters decode as nothing else either: a binary.
    if let Ok(text) = std::str::from_utf8(bytes) {
        if is_mostly_control(text) {
            return None;
        }
    }

    Some(decode_legacy(bytes, legacy))
}

/// Whether a decoded page is mostly characters no page would show. Text that is
/// mostly control characters did not come from a text file, whatever its bytes
/// happened to decode as.
fn is_mostly_control(text: &str) -> bool {
    let probe: Vec<char> = text.chars().take(8192).collect();
    if probe.is_empty() {
        return false;
    }

    let controls = probe
        .iter()
        .filter(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
        .count();

    controls * 10 > probe.len()
}

fn decode_utf16(bytes: &[u8], big_endian: bool) -> String {
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| {
            if big_endian {
                u16::from_be_bytes([pair[0], pair[1]])
            } else {
                u16::from_le_bytes([pair[0], pair[1]])
            }
        })
        .collect();

    String::from_utf16_lossy(&units)
}

/// UTF-16 with no byte order mark, read from what the bytes themselves say.
///
/// **Two shapes reach here, and they are told apart by the NUL bytes.** A NUL is the high half of
/// an ASCII character, so a UTF-16 file of mixed text carries a NUL beside every ASCII one, all on
/// the same side of the pairs — that side is the byte order, and text the other way round is
/// unreadable. NULs on *both* sides are a binary that happens to have zeroes in it.
///
/// A file with no NUL at all is a file with no ASCII at all, which only an East Asian page is
/// written like: the whole of it is CJK, kana, hangul, bopomofo or their punctuation. Such a file
/// is read little-endian, because that is the order Windows and every tool that writes `Unicode`
/// writes — and the reading has to be a page of one East Asian region, which is what keeps a GBK or
/// Shift-JIS page read as UTF-16 out of the renderer (those come out as a scatter of Hangul and
/// private-use code points, never as a page).
///
/// The decoded page is checked before it is trusted, so bytes that merely happen to look like
/// UTF-16 are still reported as the binary they are.
fn decode_utf16_without_mark(bytes: &[u8]) -> Option<String> {
    let probe = &bytes[..bytes.len().min(8192) / 2 * 2];
    if probe.len() < 4 || !bytes.len().is_multiple_of(2) {
        return None;
    }

    let mut zeros_even = 0usize;
    let mut zeros_odd = 0usize;
    for (index, byte) in probe.iter().enumerate() {
        if *byte == 0 {
            if index % 2 == 0 {
                zeros_even += 1;
            } else {
                zeros_odd += 1;
            }
        }
    }

    let zeros = zeros_even + zeros_odd;
    if zeros > 0 {
        // A NUL on both sides of the pairs is not a 16-bit encoding of text.
        if zeros_even.min(zeros_odd) * 5 > zeros {
            return None;
        }

        let big_endian = zeros_even > zeros_odd;
        let text = decode_utf16(bytes, big_endian);
        return is_a_page(&text).then_some(text);
    }

    // No NUL anywhere: no ASCII anywhere. East Asian text, and little-endian.
    let text = decode_utf16(bytes, false);
    reads_as_east_asian(&text).then_some(text)
}

/// Whether a decoded page is one a page can be made of: no control characters but the breaks and the
/// tabs, no private-use code points, no unpaired surrogates, nothing that decoded to nothing.
///
/// A `char` in Rust is never a surrogate, so what `from_utf16_lossy` made of a broken pair is the
/// replacement character — which is the one this refuses.
fn is_a_page(text: &str) -> bool {
    if text.chars().count() < 2 {
        return false;
    }

    !text.chars().any(|character| {
        let point = character as u32;
        matches!(
            point,
            0x0000..=0x0008 | 0x000B | 0x000C | 0x000E..=0x001F | 0x007F..=0x009F
                | 0xE000..=0xF8FF | 0xFFFD..=0xFFFF
        )
    })
}

/// Whether every character of a decoded page is one an East Asian page is written in — and there
/// are enough of them for it to be a page rather than a coincidence of two byte pairs.
///
/// The regions are the ones a page of Chinese, Japanese or Korean is drawn from: the CJK radicals,
/// kana, hangul, bopomofo and unified ideographs, the compatibility and fullwidth forms, and the
/// punctuation and spaces that go between them.
fn reads_as_east_asian(text: &str) -> bool {
    const EAST_ASIAN: [(u32, u32); 11] = [
        (0x2E80, 0x303F), // CJK radicals, Kangxi, CJK symbols and punctuation
        (0x3040, 0x312F), // hiragana, katakana, bopomofo
        (0x3130, 0x318F), // hangul compatibility jamo
        (0x3200, 0x33FF), // enclosed CJK, CJK compatibility
        (0x3400, 0x4DBF), // CJK extension A
        (0x4E00, 0x9FFF), // CJK unified ideographs
        (0xAC00, 0xD7AF), // hangul syllables
        (0xF900, 0xFAFF), // CJK compatibility ideographs
        (0xFE30, 0xFE4F), // CJK compatibility forms
        (0xFF01, 0xFF5E), // fullwidth forms
        (0xFF61, 0xFFDC), // halfwidth katakana
    ];

    let mut characters = 0usize;
    let mut distinct = std::collections::HashSet::new();
    for character in text.chars() {
        if matches!(character, '\t' | '\n' | '\r') {
            continue;
        }
        let point = character as u32;
        // The general punctuation a Chinese page is written with, and the ideographic space.
        let east_asian = EAST_ASIAN
            .iter()
            .any(|(low, high)| (*low..=*high).contains(&point))
            || (0x2000..=0x206F).contains(&point);
        if !east_asian {
            return false;
        }
        characters += 1;
        distinct.insert(character);
    }

    // Six characters at the least: two byte pairs are a code point, not a page.
    characters >= 6 && distinct.len() >= 4
}

/// Windows-1252 agrees with Latin-1 outside `0x80..=0x9F`, where it borrows
/// punctuation from its own table. CP437 is a different table over `0x80..=0xFF`.
fn decode_legacy(bytes: &[u8], encoding: LegacyEncoding) -> String {
    bytes
        .iter()
        .map(|byte| match encoding {
            LegacyEncoding::Windows1252 => cp1252_char(*byte),
            LegacyEncoding::Cp437 => cp437_char(*byte),
        })
        .collect()
}

fn cp1252_char(byte: u8) -> char {
    const HIGH: [char; 32] = [
        '\u{20AC}', '\u{FFFD}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{FFFD}',
        '\u{017D}', '\u{FFFD}', '\u{FFFD}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}',
        '\u{0153}', '\u{FFFD}', '\u{017E}', '\u{0178}',
    ];

    match byte {
        0x00..=0x7F => byte as char,
        0x80..=0x9F => HIGH[(byte - 0x80) as usize],
        _ => byte as char,
    }
}

/// CP437's upper half, in code point order from `0x80`.
fn cp437_char(byte: u8) -> char {
    const HIGH: [char; 128] = [
        'Ç', 'ü', 'é', 'â', 'ä', 'à', 'å', 'ç', 'ê', 'ë', 'è', 'ï', 'î', 'ì', 'Ä', 'Å', 'É', 'æ',
        'Æ', 'ô', 'ö', 'ò', 'û', 'ù', 'ÿ', 'Ö', 'Ü', '¢', '£', '¥', '₧', 'ƒ', 'á', 'í', 'ó', 'ú',
        'ñ', 'Ñ', 'ª', 'º', '¿', '⌐', '¬', '½', '¼', '¡', '«', '»', '░', '▒', '▓', '│', '┤', '╡',
        '╢', '╖', '╕', '╣', '║', '╗', '╝', '╜', '╛', '┐', '└', '┴', '┬', '├', '─', '┼', '╞', '╟',
        '╚', '╔', '╩', '╦', '╠', '═', '╬', '╧', '╨', '╤', '╥', '╙', '╘', '╒', '╓', '╫', '╪', '┘',
        '┌', '█', '▄', '▌', '▐', '▀', 'α', 'ß', 'Γ', 'π', 'Σ', 'σ', 'µ', 'τ', 'Φ', 'Θ', 'Ω', 'δ',
        '∞', 'φ', 'ε', '∩', '≡', '±', '≥', '≤', '⌠', '⌡', '÷', '≈', '°', '∙', '·', '√', 'ⁿ', '²',
        '■', '\u{00A0}',
    ];

    match byte {
        0x00..=0x7F => byte as char,
        _ => HIGH[(byte - 0x80) as usize],
    }
}

/// Tabs are replaced with spaces before anything else looks at the text, so the
/// highlighter, the Markdown parser and the layout all see the same columns, and
/// no tab character ever reaches GDI's own tab stops. Carriage returns go with
/// them, which is what makes Windows and Unix line endings read the same.
fn expand_tabs(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut column = 0usize;

    for character in text.chars() {
        match character {
            '\t' => {
                let spaces = TAB_WIDTH - (column % TAB_WIDTH);
                for _ in 0..spaces {
                    out.push(' ');
                }
                column += spaces;
            }
            '\n' => {
                out.push('\n');
                column = 0;
            }
            '\r' => {}
            _ => {
                out.push(character);
                column += 1;
            }
        }
    }

    out
}

// ----------------------------------------------------------------------- RTF

/// Destinations whose contents are not document text: font, color and style
/// tables, embedded pictures, field instructions and headers.
const RTF_IGNORED_DESTINATIONS: &[&str] = &[
    "fonttbl",
    "filetbl",
    "colortbl",
    "stylesheet",
    "listtable",
    "listoverridetable",
    "revtbl",
    "rsidtbl",
    "generator",
    "info",
    "pict",
    "object",
    "themedata",
    "datastore",
    "fldinst",
    "header",
    "headerl",
    "headerr",
    "headerf",
    "footer",
    "footerl",
    "footerr",
    "footerf",
    "footnote",
    "xmlnstbl",
];

/// Reduce RTF to the text it carries.
///
/// Word writes its tables ahead of the body, and those groups are skipped whole;
/// what is left is the document with its paragraphs, tabs, escaped bytes and
/// `\uN` sequences resolved. This is the text-stripping half of RTF and
/// deliberately not a layout engine: the preview shows the file's text, without
/// its formatting.
pub(super) fn strip_rtf(source: &str) -> Vec<String> {
    let characters: Vec<char> = source.chars().collect();
    let mut lines: Vec<String> = Vec::new();
    let mut line = String::new();
    let mut index = 0usize;
    let mut depth = 0usize;
    let mut ignoring: Option<usize> = None;
    let mut unicode_skip = 1usize;

    while index < characters.len() {
        match characters[index] {
            '{' => {
                depth += 1;
                index += 1;
            }
            '}' => {
                if ignoring.map(|start| depth <= start).unwrap_or(false) {
                    ignoring = None;
                }
                depth = depth.saturating_sub(1);
                index += 1;
            }
            '\\' => {
                index += 1;
                if index >= characters.len() {
                    break;
                }

                match characters[index] {
                    '\\' | '{' | '}' => {
                        if ignoring.is_none() {
                            line.push(characters[index]);
                        }
                        index += 1;
                    }
                    '\'' => {
                        let hex: String = characters.iter().skip(index + 1).take(2).collect();
                        if let Ok(value) = u8::from_str_radix(&hex, 16) {
                            if ignoring.is_none() {
                                line.push(cp1252_char(value));
                            }
                        }
                        index += 3;
                    }
                    '*' => {
                        ignoring = Some(depth);
                        index += 1;
                    }
                    character if character.is_ascii_alphabetic() => {
                        let mut word = String::new();
                        while index < characters.len() && characters[index].is_ascii_alphabetic() {
                            word.push(characters[index]);
                            index += 1;
                        }

                        let mut negative = false;
                        let mut parameter: Option<i64> = None;
                        if index < characters.len()
                            && (characters[index] == '-' || characters[index].is_ascii_digit())
                        {
                            if characters[index] == '-' {
                                negative = true;
                                index += 1;
                            }
                            let mut digits = String::new();
                            while index < characters.len() && characters[index].is_ascii_digit() {
                                digits.push(characters[index]);
                                index += 1;
                            }
                            let value: i64 = digits.parse().unwrap_or(0);
                            parameter = Some(if negative { -value } else { value });
                        }

                        // A trailing space belongs to the control word.
                        if index < characters.len() && characters[index] == ' ' {
                            index += 1;
                        }

                        if word == "uc" {
                            unicode_skip = parameter.unwrap_or(1).max(0) as usize;
                            continue;
                        }
                        if word == "u" {
                            if ignoring.is_none() {
                                if let Some(value) = parameter {
                                    let code = if value < 0 { value + 65536 } else { value };
                                    if let Some(character) = char::from_u32(code as u32) {
                                        line.push(character);
                                    }
                                }
                            }
                            // The escape is followed by `unicode_skip` fallback
                            // characters, which the document does not need.
                            for _ in 0..unicode_skip {
                                skip_rtf_character(&characters, &mut index);
                            }
                            continue;
                        }

                        if ignoring.is_some() {
                            continue;
                        }

                        if RTF_IGNORED_DESTINATIONS.contains(&word.as_str()) {
                            ignoring = Some(depth);
                            continue;
                        }

                        match word.as_str() {
                            "par" | "line" | "sect" => lines.push(std::mem::take(&mut line)),
                            "tab" => line.push('\t'),
                            "cell" => line.push('\t'),
                            "row" => lines.push(std::mem::take(&mut line)),
                            "bullet" => line.push('•'),
                            "emdash" => line.push('—'),
                            "endash" => line.push('–'),
                            "lquote" => line.push('‘'),
                            "rquote" => line.push('’'),
                            "ldblquote" => line.push('“'),
                            "rdblquote" => line.push('”'),
                            _ => {}
                        }
                    }
                    _ => index += 1,
                }
            }
            '\r' | '\n' => index += 1,
            character => {
                if ignoring.is_none() {
                    line.push(character);
                }
                index += 1;
            }
        }
    }

    if !line.trim().is_empty() {
        lines.push(line);
    }

    lines
        .iter()
        .map(|line| line.trim_end().to_string())
        .collect()
}

/// Consume one character of RTF source, whether it is plain text or an escape,
/// which is how `\uN`'s fallback characters are counted.
fn skip_rtf_character(characters: &[char], index: &mut usize) {
    if *index >= characters.len() {
        return;
    }

    if characters[*index] != '\\' {
        *index += 1;
        return;
    }

    *index += 1;
    if *index < characters.len() && characters[*index] == '\'' {
        *index = (*index + 3).min(characters.len());
        return;
    }

    while *index < characters.len() && !characters[*index].is_whitespace() {
        *index += 1;
    }
}

// ---------------------------------------------------------------------- ANSI

pub(super) struct AnsiPalette {
    foreground: [[u8; 3]; 16],
    background: [[u8; 3]; 16],
}

/// The console palette, pulled toward the page's own contrast: an NFO's colors
/// assume a black console, so on a light page its dark colors would disappear.
impl AnsiPalette {
    const BASE: [[u8; 3]; 16] = [
        [12, 12, 12],
        [197, 15, 31],
        [19, 161, 14],
        [193, 156, 0],
        [0, 55, 218],
        [136, 23, 152],
        [58, 150, 221],
        [204, 204, 204],
        [118, 118, 118],
        [231, 72, 86],
        [22, 198, 12],
        [249, 241, 165],
        [59, 120, 255],
        [180, 0, 158],
        [97, 214, 214],
        [242, 242, 242],
    ];

    pub(super) fn new(theme: &LoadedTheme) -> Self {
        let page = rgb(theme.background());
        let foreground = rgb(theme.foreground());

        let mut palette = Self {
            foreground: [[0, 0, 0]; 16],
            background: [[0, 0, 0]; 16],
        };

        for index in 0..16 {
            palette.foreground[index] = readable(Self::BASE[index], page);
            palette.background[index] = blend(page, Self::BASE[index], 0.25);
        }

        // Console black is the one color that has to become the page's own text
        // color to stay visible, whichever page it is on.
        palette.foreground[0] = foreground;

        palette
    }
}

pub(super) fn push_ansi_spans(
    spans: &mut Vec<Span>,
    line: &str,
    base: TextStyle,
    palette: &AnsiPalette,
) {
    let mut style = base;
    let mut pending = String::new();
    let mut characters = line.chars().peekable();

    while let Some(character) = characters.next() {
        if character != '\u{1B}' {
            pending.push(character);
            continue;
        }

        if characters.peek() != Some(&'[') {
            continue;
        }
        characters.next();

        let mut parameters = String::new();
        let mut terminator = None;
        for next in characters.by_ref() {
            if next.is_ascii_digit() || next == ';' {
                parameters.push(next);
            } else {
                terminator = Some(next);
                break;
            }
        }

        let flushes = terminator == Some('m');
        if flushes {
            if !pending.is_empty() {
                spans.push(Span {
                    text: std::mem::take(&mut pending),
                    style,
                });
            }
            apply_sgr(&mut style, &parameters, base, palette);
        } else if terminator.is_none() {
            // A truncated sequence: the rest of the line is text.
            break;
        }
    }

    if !pending.is_empty() {
        spans.push(Span {
            text: pending,
            style,
        });
    }
}

fn apply_sgr(style: &mut TextStyle, parameters: &str, base: TextStyle, palette: &AnsiPalette) {
    let values: Vec<u32> = if parameters.is_empty() {
        vec![0]
    } else {
        parameters
            .split(';')
            .map(|value| value.parse().unwrap_or(0))
            .collect()
    };

    for value in values {
        match value {
            0 => {
                style.foreground = base.foreground;
                style.background = None;
                style.bold = false;
                style.underline = false;
            }
            1 => style.bold = true,
            22 => style.bold = false,
            4 => style.underline = true,
            24 => style.underline = false,
            7 => {
                let foreground = style.foreground;
                style.foreground = style.background.unwrap_or(base.foreground);
                style.background = Some(foreground);
            }
            30..=37 => style.foreground = palette.foreground[(value - 30) as usize],
            39 => style.foreground = base.foreground,
            40..=47 => style.background = Some(palette.background[(value - 40) as usize]),
            49 => style.background = None,
            90..=97 => style.foreground = palette.foreground[(value - 90 + 8) as usize],
            100..=107 => style.background = Some(palette.background[(value - 100 + 8) as usize]),
            _ => {}
        }
    }
}
