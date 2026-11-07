//! Whether a DXF holds geometry the viewer can draw, which is the whole of what the model claim
//! asks of a drawing.
//!
//! f3d reads a DXF through Assimp's importer rather than through a reader of its own: the ASCII
//! spelling only, and only the entities that carry a mesh — a `3DFACE`, and a `POLYLINE` whose flag
//! word says it is a 3D one, a mesh or a polyface. Everything else a drafting sheet is made of
//! (`LWPOLYLINE`, `ARC`, `CIRCLE`, `TEXT`/`MTEXT`, `DIMENSION`, `HATCH`, `SPLINE`, `ELLIPSE`, a 2D
//! `POLYLINE`, and the ACIS solids — `3DSOLID`, `BODY`, `REGION`, `SURFACE` — which the importer
//! cannot read at all) draws nothing there, while the render engine beside it imports the whole
//! sheet as a drawing. So the two are told apart by what the file holds rather than by what it is
//! called, and this is that question asked of the group codes rather than of the text: a `TEXT`
//! whose own string reads `3DFACE` is a label on a drawing, and a search of the bytes would take it
//! for an entity.
//!
//! The scan is bounded twice over, by a byte count and by a pair count, because it is asked on the
//! thread that pumps a hover's window and the sections in front of `ENTITIES` can be megabytes. A
//! file whose bound is reached with no such entity in it is the render engine's, which is the
//! answer a drafting sheet — by far the common case — should pay for anyway.

use std::fs::File;
use std::io::{BufRead, BufReader, Read};
use std::path::Path;

/// How much of a drawing is read before the answer is the render engine's.
const SCAN_BYTES: u64 = 256 * 1024;

/// How many group-code pairs are parsed, which bounds a file of one enormous section a second way.
const SCAN_PAIRS: usize = 20_000;

/// The group-code 70 bits that make a `POLYLINE` one f3d imports: 3D polyline, 3D mesh, polyface
/// mesh. A 2D `POLYLINE` — the flags of a drafting sheet, and of every filled outline — draws
/// nothing in the viewer.
const POLYLINE_FLAGS_3D: u32 = 8 | 16 | 64;

/// Whether the drawing at `path` holds one of the entities the viewer imports.
///
/// A file that is not there, one written in the binary spelling — whose sentinel sits where the
/// first group code would be — and one whose bound is reached first are all answered false, which
/// is the render engine drawing the sheet.
pub(crate) fn is_3d(path: &Path) -> bool {
    let Ok(file) = File::open(path) else {
        return false;
    };
    let mut reader = BufReader::new(file.take(SCAN_BYTES));
    let mut line = Vec::new();

    // What a group code of 0 names is an entity or a structural marker, and only two of the file's
    // sections hold geometry that is drawn: the drawing's own entities, and the blocks an `INSERT`
    // places, which is where a repeated solid is written.
    let mut in_a_scanned_section = false;
    // The pair after a `SECTION` marker names the section, which is a group code of 2.
    let mut after_a_section = false;
    // A `POLYLINE` is settled by its own flag word, which follows it in the same entity.
    let mut inside_a_polyline = false;

    for _ in 0..SCAN_PAIRS {
        let Some(code) = next_code(&mut reader, &mut line) else {
            return false;
        };
        let Some(value) = next_value(&mut reader, &mut line) else {
            return false;
        };

        match code {
            0 => {
                in_a_scanned_section = in_a_scanned_section && value != "ENDSEC" && value != "EOF";
                after_a_section = value == "SECTION";
                inside_a_polyline = value == "POLYLINE";

                if in_a_scanned_section && value == "3DFACE" {
                    return true;
                }
            }
            2 if after_a_section => {
                after_a_section = false;
                in_a_scanned_section = value == "BLOCKS" || value == "ENTITIES";
            }
            70 if inside_a_polyline => {
                // One flag word per polyline, so a `70` of the entity after it is not this one's.
                inside_a_polyline = false;

                if value
                    .parse::<u32>()
                    .is_ok_and(|flags| flags & POLYLINE_FLAGS_3D != 0)
                {
                    return true;
                }
            }
            _ => {}
        }
    }

    false
}

/// The next group code, or nothing at the end of the window: a line that is not a number is not a
/// drawing's group code, which is what the binary spelling opens with instead of one.
fn next_code(reader: &mut impl BufRead, line: &mut Vec<u8>) -> Option<i32> {
    read_line(reader, line)?;

    std::str::from_utf8(line).ok()?.trim().parse().ok()
}

/// The value the group code above it belongs to, which is text of any shape — an empty one
/// included — and is never read as a code itself.
fn next_value<'a>(reader: &mut impl BufRead, line: &'a mut Vec<u8>) -> Option<&'a str> {
    read_line(reader, line)?;

    std::str::from_utf8(line).ok().map(str::trim)
}

fn read_line(reader: &mut impl BufRead, line: &mut Vec<u8>) -> Option<()> {
    line.clear();

    let read = reader.read_until(b'\n', line).ok()?;

    (read > 0).then_some(())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A drawing of one entity list, in the sections a real file carries in front of it: the
    /// header, the tables, and then whatever the caller wrote.
    fn drawing(entities: &str) -> String {
        format!(
            "0\nSECTION\n2\nHEADER\n9\n$ACADVER\n1\nAC1015\n0\nENDSEC\n\
             0\nSECTION\n2\nTABLES\n0\nENDSEC\n{entities}"
        )
    }

    fn section(name: &str, entities: &str) -> String {
        format!("0\nSECTION\n2\n{name}\n{entities}0\nENDSEC\n")
    }

    fn written(name: &str, body: &str) -> std::path::PathBuf {
        let folder = std::env::temp_dir().join("rust-hover-preview-dxf-scan");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join(name);
        std::fs::write(&path, body).expect("a drawing");

        path
    }

    fn is_3d_of(name: &str, body: &str) -> bool {
        is_3d(&written(name, body))
    }

    /// The two entities the viewer imports, in the section a drawing writes them in.
    #[test]
    fn a_face_and_a_3d_polyline_are_the_viewers_and_a_flat_one_is_not() {
        assert!(
            is_3d_of(
                "face.dxf",
                &drawing(&section("ENTITIES", "0\n3DFACE\n8\n0\n10\n0\n11\n1\n"))
            ),
            "a 3DFACE is the entity f3d's importer reads"
        );

        for (name, flags) in [("a 3d polyline", 8), ("a 3d mesh", 16), ("a polyface", 64)] {
            assert!(
                is_3d_of(
                    &format!("polyline-{flags}.dxf"),
                    &drawing(&section(
                        "ENTITIES",
                        &format!("0\nPOLYLINE\n8\n0\n66\n1\n70\n{flags}\n")
                    ))
                ),
                "`{name}` is a polyline the importer reads: 70 says so"
            );
        }

        // A 2D polyline is the outline every drafting sheet is drawn with, and the viewer draws
        // nothing of it: the entity the model claim is asked about is not the name of one.
        assert!(
            !is_3d_of(
                "flat.dxf",
                &drawing(&section("ENTITIES", "0\nPOLYLINE\n8\n0\n70\n1\n"))
            ),
            "a 2D POLYLINE is not geometry the viewer imports"
        );
    }

    /// **The scan reads group-code pairs and never the text of the file**, which is the whole
    /// reason it exists: a drawing whose label, layer or block name is spelled `3DFACE` is a
    /// drawing of a sheet, and a search of the bytes would hand it to a viewer that draws nothing.
    #[test]
    fn a_name_inside_a_value_is_not_an_entity() {
        assert!(
            !is_3d_of(
                "labelled.dxf",
                &drawing(&section(
                    "ENTITIES",
                    // The `1` and `2` group codes carry strings, and both of these read
                    // `3DFACE`: a text entity's own content, and a layer's name.
                    "0\nTEXT\n8\n3DFACE\n10\n0\n1\n3DFACE\n\
                     0\nLWPOLYLINE\n8\n3DFACE\n90\n2\n70\n0\n"
                ))
            ),
            "a `3DFACE` written as a value is a name a drawing carries"
        );
    }

    /// The geometry the viewer cannot draw: the modern 2D entities, and the ACIS solids, which are
    /// a separate kernel's bodies rather than a mesh. Each of them is a sheet the render engine
    /// imports whole.
    #[test]
    fn the_entities_the_viewer_has_no_reader_for_are_not_its() {
        assert!(
            !is_3d_of(
                "sheet.dxf",
                &drawing(&section(
                    "ENTITIES",
                    "0\nLWPOLYLINE\n90\n4\n70\n0\n\
                     0\nARC\n10\n0\n20\n0\n40\n5\n\
                     0\nCIRCLE\n10\n0\n20\n0\n40\n5\n\
                     0\nSPLINE\n70\n8\n\
                     0\nELLIPSE\n10\n0\n\
                     0\nHATCH\n2\n0\n\
                     0\nDIMENSION\n70\n32\n\
                     0\nMTEXT\n1\nnote\n\
                     0\n3DSOLID\n70\n1\n"
                ))
            ),
            "a drafting sheet of 2D entities and an ACIS body is the render engine's"
        );

        // A `SPLINE` and a `3DSOLID` carry a flag word of 8 and one of 1: neither is a polyline,
        // and the two are why the flag word is only read inside one.
        assert!(
            !is_3d_of(
                "acis.dxf",
                &drawing(&section("ENTITIES", "0\nSPLINE\n70\n8\n0\nBODY\n70\n64\n"))
            ),
            "the flag word of an entity that is not a POLYLINE settles nothing"
        );
    }

    /// Both of the sections a solid is written in are scanned: a block is where a repeated one
    /// lives, and a drawing whose only 3D geometry is a block the sheet has not placed is still a
    /// file the viewer can draw.
    #[test]
    fn the_geometry_a_block_holds_is_the_viewers_too() {
        assert!(
            is_3d_of(
                "blocks.dxf",
                &drawing(&section(
                    "BLOCKS",
                    "0\nBLOCK\n2\nSOLID\n0\n3DFACE\n8\n0\n0\nENDBLK\n"
                ))
            ),
            "a block's own entities are geometry the viewer imports"
        );
    }

    /// The bound, both ways: a marker past it is the render engine's answer, and the binary
    /// spelling — whose sentinel sits where the first group code would be — is one the viewer
    /// cannot read at all.
    #[test]
    fn a_drawing_past_the_bound_and_one_written_as_binary_are_the_render_engines() {
        let mut late = drawing(&section("ENTITIES", "0\nLWPOLYLINE\n90\n4\n"));
        // Padded with entities in a section of its own, so the cap is what stops the scan rather
        // than the end of the file.
        let padding = "0\nTEXT\n1\nfiller\n".repeat(SCAN_BYTES as usize / 16);
        late.push_str(&section("ENTITIES", &padding));
        late.push_str(&section("ENTITIES", "0\n3DFACE\n8\n0\n"));
        assert!(
            !is_3d_of("late.dxf", &late),
            "a drawing whose geometry is past the bound is the render engine's"
        );

        assert!(
            !is_3d_of(
                "binary.dxf",
                "AutoCAD Binary DXF\r\n\u{1a}\u{0}\u{0}0\u{0}S\u{0}E\u{0}C\u{0}T\u{0}"
            ),
            "the binary spelling is not one the viewer's importer reads"
        );
    }

    /// **A drawing is written with its group codes right-justified**, so the scanner is handed
    /// `  0` in front of every one of them: what it reads is the code a line names rather than the
    /// column it starts in, which is why the padding in front of a front is the front's question
    /// alone (see `ascii_dxf`).
    #[test]
    fn a_code_written_in_the_columns_a_drawing_is_written_with_is_still_a_code() {
        assert!(
            is_3d_of(
                "padded.dxf",
                "  0\r\nSECTION\r\n  2\r\nENTITIES\r\n  0\r\n3DFACE\r\n  8\r\n0\r\n  0\r\nENDSEC\r\n"
            ),
            "a padded `3DFACE` is the entity the importer reads"
        );

        assert!(
            !is_3d_of(
                "padded-flat.dxf",
                "  0\r\nSECTION\r\n  2\r\nENTITIES\r\n  0\r\nPOLYLINE\r\n 70\r\n1\r\n"
            ),
            "and a padded 2D polyline is still the sheet's"
        );
    }

    /// A file that is not there, an empty one, and one whose first line is not a group code are
    /// all answered false: nothing is opened on their account twice.
    #[test]
    fn nothing_to_read_is_not_geometry() {
        assert!(!is_3d(Path::new("no-such-drawing.dxf")));
        assert!(!is_3d_of("empty.dxf", ""));
        assert!(!is_3d_of("truncated.dxf", "0\nSECTION\n2\n"));
    }
}
