//! What a model's own head says about the model its name promises.
//!
//! A name in `[f3d]` used to be the whole of the claim: a `.bin` was a point cloud because it was
//! called one, and a file that was not went to the viewer anyway, which spent a launch and a black
//! window finding out (see `f3d_render`). What is asked instead is the front of the file — the one
//! window every other table of this app is answered from (see `head`) — and a file that fails the
//! question is not a model at all: the route answers with no kind, so nothing is drawn and nothing
//! is started, which is the answer a name outside every list already had.
//!
//! Most of the names have a magic, and it is the format's own rather than a guess: `LASF` opens a
//! LAS, `BPF!` a BPF, `CSD` an Optech cloud, the dated version word of its header a TerraScan BIN,
//! `Ogawa` an Alembic, `# vtk DataFile` a legacy VTK, `PXR-USDC` a crate USD, `Kaydara FBX Binary`
//! a binary FBX, and
//! `ISO-10303-21;` is the header of every STEP document there is.
//! Four families are declared by the *box* the file arrives in rather than by the format: a
//! `.3mf`, a `.fcstd`, a `.slpk` and a `.usdz` are zips, a `.nc`, a `.vtkhdf` and an `.abc` may be
//! HDF5, and each is accepted on the box alone, because what is inside it is a question its own
//! reader asks afterwards.
//!
//! The rest are text whose first line names them — `solid`, `ply`, `OFF`, `NRRD`, `#VRML`,
//! `#?RADIANCE`, the section letter an IGES writes in column 73, an OBJ statement at the start of a
//! line — and the check is as strong as such a check can be, which is to say it tells a model
//! from a document and does not pretend to read the document. A header is read as far as the file
//! stays text: a `.vtk`, a `.vtp`, a `.brep`, a `.mha` and a `.pic` are a line naming the format in
//! front of a payload that is not text at all, and the payload is no reason to refuse the header. A `.pic` is worth naming here: f3d's
//! reader for it is the Radiance HDR one rather than the Softimage PIC the extension also stands
//! for, so `#?RADIANCE` is what its head is asked for.
//!
//! Four names declare nothing in their front at all — `sbet`, a trajectory of binary records;
//! `qi`, the QFIT cloud whose header holds nothing this side can check; `splat`, whose
//! files are either PLY or a bare run of gaussian records; and `xbf`, an OpenCASCADE cache — so
//! nothing is read of a file of one of those names and the name is the whole of the claim: the
//! viewer is started, and it is the one that answers, exactly as it was before this module
//! existed.
//!
//! Two answers are deliberately the loose ones. **A name a user adds to `[f3d]` themselves is not
//! in the table and is trusted**: nothing here knows the format it stands for, so nothing here can
//! refuse it. And where the engine's own check is stricter than this one — 4.0.1 asks a `.step`
//! for a `FILE_SCHEMA`, a `.glb` for a length that agrees with the file, a `.splat` for PLY — this
//! is the weaker half on purpose: a preview refused is a preview lost, while a launch the viewer
//! turns down is the one it already paid.
//!
//! A front that cannot be read at all — a file on a cloud volume, a file that has just gone — is
//! trusted too, for the same reason: what the viewer says about it is the answer, and it is the
//! one it gave before.
//!
//! **One name is asked of its body rather than of its front.** A DXF is a drawing, and the viewer
//! reads only the 3D half of the format — the meshes — while the render engine beside it imports the
//! 2D sheet whole; the two are told apart by which entities the file holds, which the front says
//! nothing about (see `dxf`). It is also the one name of this list a second kind can draw, so what
//! the routing does with a file of it is asked of both readers (see `routing::claim_models`).

use crate::formats::head::{self, Facts};
use std::path::Path;

mod dxf;

/// Whether the front of `path` is the model its own name says it is.
///
/// A file whose name carries no extension, or one the table does not know, answers true: the list
/// that named it is the user's own, and this module has no format to hold it to.
pub(crate) fn claims_content(path: &Path, facts: Option<&Facts>) -> bool {
    let Some(extension) = extension_of(path) else {
        return true;
    };
    let Some((_, matches)) = MODEL_FORMATS.iter().find(|(name, _)| *name == extension) else {
        return true;
    };
    // Answered before a head is read: nothing in the front of a format that declares nothing could
    // say whether the file is that format, so no file is opened on the four names' account.
    let Some(matches) = matches else {
        return true;
    };

    let head = match facts {
        Some(facts) => head::of_with_facts(path, facts),
        None => head::of(path),
    };
    // A front that cannot be read at all is trusted, since what the viewer says about it is the
    // answer — except a drawing's, whose question is its body rather than its front: a file that is
    // not there holds no geometry, and the render engine has the sheet.
    let Some(head) = head else {
        return !is_a_dxf(path);
    };
    // Asked of the entry the walk has already read where there is one, and of the file only where
    // there is none: two of the matchers need a length, and nothing else does.
    let len = facts.map_or_else(
        || std::fs::metadata(path).map_or(0, |metadata| metadata.len()),
        Facts::len,
    );

    if !matches(head.bytes(), len) {
        return false;
    }

    // A DXF's front says the file is one and nothing about whether it is a three-dimensional one,
    // which is the only half of the format the viewer draws: its body is scanned for the entities
    // the importer reads (see `dxf`). Answering the front is the table's — a matcher is handed a
    // window — and the body is asked beside it.
    extension != "dxf" || dxf::is_3d(path)
}

/// Whether `path` is a DXF, which is the one name of the model list a second kind can draw as well:
/// the render engine imports the 2D sheet the viewer draws nothing of (see `dxf` and
/// `routing::claim_libre`).
pub(crate) fn is_a_dxf(path: &Path) -> bool {
    extension_of(path).is_some_and(|extension| extension == "dxf")
}

/// The lowercased extension `path` carries, with no dot in front of it.
fn extension_of(path: &Path) -> Option<String> {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(str::to_ascii_lowercase)
}

/// Whether the front of a file, and the length of it, are the model a name promises.
type Matcher = fn(&[u8], u64) -> bool;

/// What each model name's front has to say, by the extension the name carries.
///
/// The table is the whole of the promise: a name in `[f3d]` and not here is a name a user added,
/// and is trusted. Two names may share one matcher, and often do — a `.stp`, a `.p21`, a `.210`
/// and a `.step` are one format — and the test below holds the rows and
/// `DEFAULT_F3D_EXTENSIONS` to each other in both directions.
const MODEL_FORMATS: &[(&str, Option<Matcher>)] = &[
    ("210", Some(step)),
    ("3ds", Some(three_ds)),
    ("3mf", Some(zip)),
    ("abc", Some(ogawa_or_hdf5)),
    ("amf", Some(amf_or_zip)),
    ("bin", Some(terrasolid)),
    ("bpf", Some(bpf)),
    ("brep", Some(brep)),
    ("cdf", Some(netcdf_or_hdf5)),
    ("csd", Some(optech)),
    ("dae", Some(collada)),
    // The front of a drawing, which is the section code every ASCII one opens with; what makes it
    // a model is asked of its body afterwards (see `is_a_dxf`).
    ("dxf", Some(ascii_dxf)),
    ("ex2", Some(netcdf_or_hdf5)),
    ("exo", Some(netcdf_or_hdf5)),
    ("fbx", Some(fbx)),
    ("fcstd", Some(zip)),
    ("glb", Some(gltf_binary)),
    ("gltf", Some(gltf_json)),
    ("gml", Some(gml)),
    ("ifc", Some(ifc)),
    ("iges", Some(iges)),
    ("igs", Some(iges)),
    ("las", Some(las)),
    ("laz", Some(las)),
    ("mdl", Some(quake_model)),
    ("mha", Some(meta_image)),
    ("mhd", Some(meta_image)),
    ("nc", Some(netcdf_or_hdf5)),
    ("ncdf", Some(netcdf_or_hdf5)),
    ("nhdr", Some(nrrd)),
    ("nrrd", Some(nrrd)),
    ("obj", Some(wavefront)),
    ("off", Some(object_file)),
    ("p21", Some(step)),
    ("pic", Some(radiance)),
    ("ply", Some(ply)),
    ("pts", Some(pts)),
    ("ptx", Some(ptx)),
    ("qi", None),
    ("sbet", None),
    ("slpk", Some(zip)),
    ("splat", None),
    ("spz", Some(gzip)),
    ("step", Some(step)),
    ("stl", Some(stl)),
    ("stp", Some(step)),
    ("stpnc", Some(step)),
    ("usd", Some(usd)),
    ("usda", Some(usda)),
    ("usdc", Some(usdc)),
    ("usdz", Some(zip)),
    ("vdb", Some(openvdb)),
    ("vrml", Some(vrml)),
    ("vti", Some(xml_vtk)),
    ("vtk", Some(vtk_legacy)),
    ("vtkhdf", Some(hdf5_or_xml_vtk)),
    ("vtm", Some(xml_vtk)),
    ("vtp", Some(xml_vtk)),
    ("vtr", Some(xml_vtk)),
    ("vts", Some(xml_vtk)),
    ("vtu", Some(xml_vtk)),
    ("wrl", Some(vrml)),
    ("xbf", None),
];

/// The front of a text file, where it reads as text at all.
///
/// A window that ends inside a character is text nonetheless, so only a sequence that is wrong in
/// the middle of itself refuses the file: half an em dash at the end of the probe is not a binary.
fn as_text(front: &[u8]) -> Option<&str> {
    let text = match std::str::from_utf8(front) {
        Ok(text) => text,
        Err(error) if error.error_len().is_none() => {
            std::str::from_utf8(&front[..error.valid_up_to()]).ok()?
        }
        Err(_) => return None,
    };

    text.chars()
        .all(|character| character.is_whitespace() || !character.is_control())
        .then_some(text)
}

/// The same, with any byte-order mark and leading whitespace taken off the front.
fn text_start(front: &[u8]) -> Option<&str> {
    let text = as_text(front)?;

    Some(text.trim_start_matches('\u{feff}').trim_start())
}

/// The text a front opens with, up to the first byte that is not text, taken off any byte-order
/// mark and leading whitespace. A header in front of a payload — a VTK's arrays, a BRep's curves,
/// a Radiance picture's pixels — is still the format's own header, and `as_text` would refuse it
/// for the bytes behind it. An empty prefix is no text at all.
fn text_head(front: &[u8]) -> Option<&str> {
    let text = match std::str::from_utf8(front) {
        Ok(text) => text,
        Err(error) => std::str::from_utf8(&front[..error.valid_up_to()]).ok()?,
    };
    let cut = text
        .find(|character: char| character.is_control() && !character.is_whitespace())
        .unwrap_or(text.len());

    let head = text[..cut].trim_start_matches('\u{feff}').trim_start();

    (!head.is_empty()).then_some(head)
}

/// The three signatures a zip opens with: a local file, a central directory, a spanned archive.
const ZIP_MAGICS: [&[u8]; 3] = [b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"];

fn zip(front: &[u8], _len: u64) -> bool {
    ZIP_MAGICS.iter().any(|magic| front.starts_with(magic))
}

/// The columns a drawing's first group code may be written in ahead of the code itself: verbatim
/// DXF right-justifies every code in three, so a file written by the application opens `  0`.
const DXF_CODE_PAD: usize = 4;

/// An AutoCAD interchange drawing in the ASCII spelling the viewer's importer reads, whatever
/// spacing in front of its first group code. The binary one opens with a sentinel of its own, and
/// is the render engine's alone.
///
/// This is the same question `content_type::matchers::is_dxf` asks, and it asks it here rather than
/// spelling it a second time: a front this accepts is a drawing on both roads to the file.
pub(crate) fn ascii_dxf(front: &[u8], _len: u64) -> bool {
    let front = front.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(front);
    let pad = front
        .iter()
        .take(DXF_CODE_PAD)
        .take_while(|byte| **byte == b' ' || **byte == b'\t')
        .count();
    let front = &front[pad..];

    front.starts_with(b"0\r\nSECTION") || front.starts_with(b"0\nSECTION")
}

/// The eight bytes every HDF5 file opens with.
fn hdf5(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"\x89HDF\r\n\x1a\n")
}

/// NetCDF classic in either version, or the HDF5 its fourth version is written in.
fn netcdf_or_hdf5(front: &[u8], len: u64) -> bool {
    hdf5(front, len) || front.starts_with(b"CDF\x01") || front.starts_with(b"CDF\x02")
}

/// Alembic's own serialisation, or the HDF5 one it also offers.
fn ogawa_or_hdf5(front: &[u8], len: u64) -> bool {
    front.starts_with(b"Ogawa") || hdf5(front, len)
}

/// A gzip stream, which is how a VRML and a compressed gaussian splat both arrive.
fn gzip(front: &[u8], _len: u64) -> bool {
    front.starts_with(&[0x1F, 0x8B])
}

/// The two words PDAL's TerraScan reader checks of the header it parses: a dated `HdrVersion`
/// (`TERRASOLID_Format_1`/`_2` of the reader's own struct, at offset 4) and the `RecogVal` 970401
/// that names the format (offset 8). `HdrSize` before them is not checked because the reader does
/// not check it.
fn terrasolid(front: &[u8], _len: u64) -> bool {
    const TERRASOLID_VERSIONS: [u32; 2] = [20_010_712, 20_020_715];
    if front.len() < 12 {
        return false;
    }
    let word =
        |at: usize| u32::from_le_bytes([front[at], front[at + 1], front[at + 2], front[at + 3]]);

    TERRASOLID_VERSIONS.contains(&word(4)) && word(8) == 970_401
}

/// The four characters every LAS and LAZ opens with, compressed or not.
fn las(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"LASF")
}

/// The BPF point cloud's own four characters.
fn bpf(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"BPF!")
}

/// PDAL's Optech reader checks this signature and reads the header behind it.
fn optech(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"CSD")
}

/// The three identifiers Quake and Half-Life models open with: two meshes and a studio model.
const QUAKE_MAGICS: [&[u8]; 3] = [b"IDPO", b"IDP2", b"IDST"];

fn quake_model(front: &[u8], _len: u64) -> bool {
    QUAKE_MAGICS.iter().any(|magic| front.starts_with(magic))
}

/// A 3D Studio file opens with its first chunk, whose id is `0x4D4D`.
fn three_ds(front: &[u8], _len: u64) -> bool {
    front.starts_with(&[0x4D, 0x4D])
}

/// Binary FBX names itself, and a text FBX mentions itself in the comment it opens with.
fn fbx(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"Kaydara FBX Binary")
        || as_text(front).is_some_and(|text| text.contains("FBX"))
}

/// A binary glTF's four-byte magic, which the version and the length follow.
fn gltf_binary(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"glTF")
}

/// A `.gltf` is JSON, and `"asset"` is the one member every one of them carries.
fn gltf_json(front: &[u8], _len: u64) -> bool {
    as_text(front).is_some_and(|text| text.contains("\"asset\""))
}

/// A crate USD names itself, and a text USD opens with the `#usda` of its first line.
fn usdc(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"PXR-USDC")
}

fn usda(front: &[u8], _len: u64) -> bool {
    text_start(front).is_some_and(|text| text.starts_with("#usda"))
}

/// A `.usd` may be either of them, and it may be the zip a `.usdz` is as well.
fn usd(front: &[u8], len: u64) -> bool {
    usda(front, len) || usdc(front, len) || zip(front, len)
}

/// A VRML document opens with the line that names its version, and a gzipped one with the stream.
fn vrml(front: &[u8], len: u64) -> bool {
    text_start(front).is_some_and(|text| text.starts_with("#VRML")) || gzip(front, len)
}

/// A legacy VTK opens with the line that names the format and its version.
fn vtk_legacy(front: &[u8], _len: u64) -> bool {
    text_head(front).is_some_and(|text| text.starts_with("# vtk DataFile"))
}

/// The XML VTK family declares its root element after the declaration it may or may not carry.
fn xml_vtk(front: &[u8], _len: u64) -> bool {
    text_head(front).is_some_and(|text| text.contains("<VTKFile"))
}

/// A `.vtkhdf` is the XML family's HDF5 spelling.
fn hdf5_or_xml_vtk(front: &[u8], len: u64) -> bool {
    hdf5(front, len) || xml_vtk(front, len)
}

/// NRRD writes its version into the magic, and a detached header carries the same one.
fn nrrd(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"NRRD")
}

/// A MetaImage header is a list of `key = value` lines, and these two are in every one of them.
fn meta_image(front: &[u8], _len: u64) -> bool {
    text_head(front).is_some_and(|text| text.contains("ObjectType") || text.contains("NDims"))
}

/// OpenVDB's magic: a space, the letters every file of the family has, and four zero bytes.
fn openvdb(front: &[u8], _len: u64) -> bool {
    front.starts_with(&[0x20, b'B', b'D', b'V', 0, 0, 0, 0])
}

/// A COLLADA document is XML whose root element is the format's name.
fn collada(front: &[u8], _len: u64) -> bool {
    as_text(front).is_some_and(|text| text.contains("<COLLADA"))
}

/// CityGML is the same shape, and a `gml` element is what its documents are made of — named either
/// by an element or by the namespace its root declares, since the first `<gml:` of a CityGML may
/// sit past the head window behind the comment and the root's own attributes.
fn gml(front: &[u8], _len: u64) -> bool {
    text_head(front).is_some_and(|text| text.contains("<gml") || text.contains("opengis.net/gml"))
}

/// An AMF document is XML whose root is `<amf`, and a zip is what a compressed one arrives in.
fn amf_or_zip(front: &[u8], len: u64) -> bool {
    as_text(front).is_some_and(|text| text.contains("<amf")) || zip(front, len)
}

/// A STEP document opens with the header of its own serialisation, and an IFC is one of those.
fn step(front: &[u8], _len: u64) -> bool {
    text_start(front).is_some_and(|text| text.starts_with("ISO-10303-21;"))
}

/// An IFC may also be the XML spelling, or the zip an IfcZIP is.
fn ifc(front: &[u8], len: u64) -> bool {
    step(front, len) || as_text(front).is_some_and(|text| text.contains("<ifc")) || zip(front, len)
}

/// IGES is written in fixed columns, and the section letter of a record is the seventy-third.
fn iges(front: &[u8], _len: u64) -> bool {
    as_text(front).is_some() && front.get(72) == Some(&b'S')
}

/// A BRep is OpenCASCADE's own dump, which names itself on the first line it writes.
fn brep(front: &[u8], _len: u64) -> bool {
    text_head(front).is_some_and(|text| {
        text.contains("DBRep_DrawableShape") || text.contains("CASCADE Topology")
    })
}

/// A Radiance picture opens by naming the format in its first line.
fn radiance(front: &[u8], _len: u64) -> bool {
    text_head(front).is_some_and(|text| text.starts_with("#?RADIANCE"))
}

/// An object file's own name is the first word, which a compressed or a negated spelling extends.
fn object_file(front: &[u8], _len: u64) -> bool {
    text_start(front).is_some_and(|text| {
        text.split_ascii_whitespace()
            .next()
            .is_some_and(|token| token.ends_with("OFF"))
    })
}

/// The statements a Wavefront OBJ is made of, each of them at the start of a line.
const OBJ_STATEMENTS: [&str; 10] = ["v", "vn", "vt", "f", "l", "p", "o", "g", "usemtl", "mtllib"];

fn wavefront(front: &[u8], _len: u64) -> bool {
    as_text(front).is_some_and(|text| {
        text.lines().any(|line| {
            line.split_ascii_whitespace()
                .next()
                .is_some_and(|word| OBJ_STATEMENTS.contains(&word))
        })
    })
}

/// A PLY opens with its own three letters and the line ending the file uses.
fn ply(front: &[u8], _len: u64) -> bool {
    front.starts_with(b"ply\n") || front.starts_with(b"ply\r\n")
}

/// A PTS cloud is a count on the first line and its points behind it.
fn pts(front: &[u8], _len: u64) -> bool {
    text_start(front).is_some_and(|text| {
        text.split_ascii_whitespace()
            .next()
            .is_some_and(|token| token.parse::<u64>().is_ok())
    })
}

/// A PTX cloud's header is text and PDAL holds it to nothing beyond that.
fn ptx(front: &[u8], _len: u64) -> bool {
    as_text(front).is_some()
}

/// A triangle mesh is text that begins `solid`, or the binary form whose declared count fits the
/// file: eighty header bytes, a triangle count, and fifty bytes a triangle. A count that fits is
/// taken even where bytes follow the last triangle, because some pipelines leave them there and a
/// preview refused is a preview lost; a count the file has no room for is not a mesh at all.
fn stl(front: &[u8], len: u64) -> bool {
    if text_start(front).is_some_and(|text| text.starts_with("solid")) {
        return true;
    }
    if front.len() < 84 {
        return false;
    }

    let triangles = u64::from(u32::from_le_bytes([
        front[80], front[81], front[82], front[83],
    ]));

    len >= 84 + 50 * triangles
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formats::lists::DEFAULT_F3D_EXTENSIONS;

    /// The names the built-in `[f3d]` list carries.
    fn names() -> Vec<&'static str> {
        DEFAULT_F3D_EXTENSIONS.split(',').collect()
    }

    /// The matcher a name is held to. The four names that declare nothing in their front have none,
    /// and are asked through `claims_content` instead.
    fn matcher(name: &str) -> Matcher {
        MODEL_FORMATS
            .iter()
            .find(|(known, _)| *known == name)
            .expect("the model list's names all have a row")
            .1
            .expect("the name asked of declares a front of its own")
    }

    /// Four kilobytes of junk no model format's front can be: every byte value once, sixteen
    /// times, so it is binary, no magic begins it, and the triangle count its offset 80 is read at
    /// does not fit in it.
    fn junk() -> Vec<u8> {
        (0..4096).map(|index| (index % 256) as u8).collect()
    }

    /// What a file of a name opens with, and a head that is not one — checked through the table, so
    /// a name that is added to it without a matcher fails here rather than in a hover.
    #[test]
    fn every_family_of_model_is_claimed_by_its_own_head_and_nothing_else_is() {
        let heads: [(&str, &[u8]); 22] = [
            ("3mf", b"PK\x03\x04\x14\x00\x00\x00\x08\x00"),
            ("nc", b"\x89HDF\r\n\x1a\n\x00\x00\x00\x00"),
            ("cdf", b"CDF\x02\x00\x00\x00\x00"),
            ("exo", b"CDF\x01\x00\x00\x00\x00"),
            ("bin", b"8\x00\x00\x00\xeb\x7d\x31\x01\xa1\xce\x0e\x00"),
            ("las", b"LASF\x00\x00\x00\x00"),
            ("laz", b"LASF\x00\x00\x00\x00"),
            ("bpf", b"BPF!\x00\x00\x00\x00"),
            ("abc", b"Ogawa\x00\x00\x00"),
            ("csd", b"CSD\x00\x00\x00\x00\x00"),
            ("3ds", b"\x4d\x4d\x00\x00\x02\x00"),
            ("fbx", b"Kaydara FBX Binary  \x00\x1a\x00"),
            ("glb", b"glTF\x02\x00\x00\x00"),
            ("gltf", b"{\n  \"asset\": {\"version\": \"2.0\"}\n}\n"),
            ("usdc", b"PXR-USDC\x00\x00\x00\x00"),
            ("usda", b"#usda 1.0\n(\n)\n"),
            ("vdb", b"\x20BDV\x00\x00\x00\x00"),
            ("vrml", b"#VRML V2.0 utf8\nShape { }\n"),
            ("vtk", b"# vtk DataFile Version 3.0\nASCII\n"),
            (
                "vtu",
                b"<?xml version=\"1.0\"?>\n<VTKFile type=\"UnstructuredGrid\">\n",
            ),
            ("nrrd", b"NRRD0004\ntype: float\n"),
            ("mhd", b"ObjectType = Image\nNDims = 3\n"),
        ];

        for (name, head) in heads {
            assert!(
                matcher(name)(head, head.len() as u64),
                "`{name}` is the format its own head declares"
            );
            assert!(
                !matcher(name)(&junk(), 4096),
                "and junk of that name is not, so no viewer is started for it"
            );
        }

        // The text families that are one line of a document rather than a magic of their own.
        for (name, head) in [
            ("stl", "solid cube\nendsolid cube\n"),
            ("ply", "ply\nformat ascii 1.0\n"),
            ("obj", "v 0 0 0\nf 1 2 3\n"),
            ("step", "ISO-10303-21;\nFILE_SCHEMA(('IFC4'));\n"),
            ("stp", "\u{feff}ISO-10303-21;\nHEADER;\n"),
            ("ifc", "ISO-10303-21;\nFILE_SCHEMA(('IFC4'));\n"),
            ("brep", "DBRep_DrawableShape\nCASCADE Topology V1\n"),
            ("off", "OFF\n3 1 0\n"),
            ("pts", "3\n0 0 0\n1 0 0\n0 1 0\n"),
            ("ptx", "1\n1\n1\n0\n0\n0\n0\n"),
            ("dae", "<?xml version=\"1.0\"?>\n<COLLADA xmlns=\"x\">\n"),
            ("gml", "<?xml version=\"1.0\"?>\n<gml:FeatureCollection>\n"),
            (
                "amf",
                "<?xml version=\"1.0\"?>\n<amf unit=\"millimeter\">\n",
            ),
            ("pic", "#?RADIANCE\nFORMAT=32-bit_rle_rgbe\n"),
        ] {
            assert!(
                matcher(name)(head.as_bytes(), head.len() as u64),
                "`{name}` is text of the shape its format declares"
            );
            assert!(
                !matcher(name)(&junk(), 4096),
                "and junk of that name is not"
            );
        }

        // An IGES is the one format whose signature is a column rather than a string.
        let mut record = vec![b' '; 72];
        record.extend_from_slice(b"S      1\n");
        assert!(matcher("iges")(&record, record.len() as u64));
        assert!(matcher("igs")(&record, record.len() as u64));
        assert!(!matcher("iges")(&junk(), 4096));
    }

    /// A header in front of a payload is still the header its format declares: the arrays, the
    /// curves and the pixels behind it are not text, and a window that carries them is refused by
    /// `as_text` alone.
    #[test]
    fn a_text_header_over_a_binary_payload_is_still_the_format_it_names() {
        let heads: [(&str, &[u8]); 6] = [
            (
                "vtk",
                b"# vtk DataFile Version 4.2\nvtk output\nBINARY\n\x00\x01\x02\xff",
            ),
            (
                "vtp",
                b"<VTKFile type=\"PolyData\">\n<AppendedData>\x00\x01\xff",
            ),
            ("vtu", b"<VTKFile type=\"UnstructuredGrid\">\n\x01\x00\xff"),
            (
                "brep",
                b"\nOpen CASCADE Topology V4, (c) Open Cascade\n\x01\x00\xff",
            ),
            ("pic", b"#?RADIANCE\nFORMAT=32-bit_rle_rgbe\n\x02\x02\xff"),
            (
                "gml",
                b"<?xml version=\"1.0\"?>\n<CityModel xmlns:gml=\"http://www.opengis.net/gml\">\x01\xff",
            ),
        ];

        for (name, head) in heads {
            assert!(
                matcher(name)(head, head.len() as u64),
                "`{name}` is the header its format declares, whatever follows it"
            );
        }

        // An empty prefix is no text at all, so a picture of one of these names is not a header.
        assert!(
            !matcher("vtk")(b"\x89PNG\r\n\x1a\n\x00\x01", 8),
            "a file that opens with bytes that are not text has no header to read"
        );
    }

    /// The names whose format declares nothing in its front: no head is read of one, so its name is
    /// the whole of the claim and the viewer is started — a text file, junk, and a path with nothing
    /// at it alike.
    #[test]
    fn the_names_whose_format_declares_nothing_are_claimed_by_their_name_alone() {
        let folder = std::env::temp_dir().join("rust-hover-preview-model-formats-frontless");
        std::fs::create_dir_all(&folder).expect("a test folder");

        for name in ["qi", "sbet", "splat", "xbf"] {
            let text = folder.join(format!("a-note.{name}"));
            std::fs::write(&text, b"just a note about the weather\n").expect("a text file");
            assert!(
                claims_content(&text, None),
                "`{name}` declares nothing in its front, so a text file of that name is claimed"
            );

            let binary = folder.join(format!("junk.{name}"));
            std::fs::write(&binary, junk()).expect("a binary file");
            assert!(
                claims_content(&binary, None),
                "and so is a binary file of that name"
            );

            // No head is asked for, so a path with nothing at it is answered all the same.
            assert!(
                claims_content(&folder.join(format!("no-such-cloud.{name}")), None),
                "a name that declares nothing is answered without reading anything"
            );
        }

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// The count rule the binary spelling of a triangle mesh is held to: the count has to fit the
    /// file rather than agree with it to the byte, since bytes after the last triangle are still
    /// the mesh it declares.
    #[test]
    fn a_binary_mesh_is_a_count_that_fits_the_file() {
        let mut mesh = vec![0u8; 80];
        mesh.extend_from_slice(&7u32.to_le_bytes());
        mesh.extend_from_slice(&[0; 50 * 7]);

        assert!(
            matcher("stl")(&mesh, mesh.len() as u64),
            "eighty bytes, a count, and fifty bytes a triangle is the format"
        );

        let mut trailing = mesh.clone();
        trailing.extend_from_slice(&[0; 4]);
        assert!(
            matcher("stl")(&trailing, trailing.len() as u64),
            "trailing bytes after the last triangle are still the mesh"
        );

        let mut overshoot = mesh.clone();
        overshoot[80..84].copy_from_slice(&8u32.to_le_bytes());
        assert!(
            !matcher("stl")(&overshoot, overshoot.len() as u64),
            "a count the file has no room for is not a mesh"
        );
    }

    /// The file the report was about: a `.bin` whose bytes are not a point cloud, and the two
    /// fronts that are. A picture of that name is not a model either — it is still routed by
    /// `content_type` as the picture it is, which is a different question asked before this one.
    #[test]
    fn a_bin_is_a_model_only_where_pdal_finds_its_identifier() {
        let folder = std::env::temp_dir().join("rust-hover-preview-model-formats-bin");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let junk_file = folder.join("junk.bin");
        std::fs::write(&junk_file, junk()).expect("a junk file");
        assert!(
            !claims_content(&junk_file, None),
            "a `.bin` of junk is not a model, so the viewer is never started for it"
        );

        // The header of a real file: 56 bytes, version 20020715, identifier 970401.
        let mut head = b"8\x00\x00\x00\xeb\x7d\x31\x01\xa1\xce\x0e\x00".to_vec();
        head.extend_from_slice(&[0; 64]);
        let cloud = folder.join("cloud.bin");
        std::fs::write(&cloud, &head).expect("a cloud file");
        assert!(
            claims_content(&cloud, None),
            "the two words PDAL's reader checks are what a TerraScan BIN opens with"
        );

        // The identifier where the old check looked for it, and a version word there that is not
        // one of the two dated ones: the reader asks both, so this is not the format.
        let mut misplaced = 970_401u32.to_le_bytes().to_vec();
        misplaced.extend_from_slice(&[0; 64]);
        assert!(
            !matcher("bin")(&misplaced, misplaced.len() as u64),
            "970401 at four bytes in is not the header PDAL parses"
        );

        let mut unidentified = 56u32.to_le_bytes().to_vec();
        unidentified.extend_from_slice(&20_020_715u32.to_le_bytes());
        unidentified.extend_from_slice(&[0; 64]);
        assert!(
            !matcher("bin")(&unidentified, unidentified.len() as u64),
            "a dated version with a recog value of zero is not a TerraScan header either"
        );

        let mut picture = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
        picture.extend_from_slice(&[0; 64]);
        assert!(
            !matcher("bin")(&picture, picture.len() as u64),
            "a picture named `.bin` is not a model: it is routed as the picture its head says"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// A name the user added to `[f3d]` themselves is trusted, because nothing here knows the
    /// format it stands for; a mesh is claimed, and a file that is neither has no kind.
    #[test]
    fn a_name_of_the_users_own_is_trusted_and_a_mesh_is_claimed_by_its_own_head() {
        let folder = std::env::temp_dir().join("rust-hover-preview-model-formats-names");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let added = folder.join("shape.xyz");
        std::fs::write(&added, junk()).expect("a file of a name of the user's own");
        assert!(
            claims_content(&added, None),
            "an extension this table has never seen is a name the user vouched for"
        );

        let mesh = folder.join("mesh.stl");
        std::fs::write(&mesh, b"solid cube\nendsolid cube\n").expect("an ascii mesh");
        assert!(
            claims_content(&mesh, None),
            "an ascii mesh is the format its name says"
        );

        let renamed = folder.join("renamed.stl");
        std::fs::write(&renamed, junk()).expect("a file that is not a mesh");
        assert!(
            !claims_content(&renamed, None),
            "and a file of that name which is neither solid nor a mesh whose count fits it is not"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// A drawing is the one name whose body is asked as well as its front: the viewer imports the
    /// 3D half of the format and nothing of the 2D sheet, so the geometry the file holds is what
    /// the claim is answered by (see `dxf`).
    #[test]
    fn a_drawing_is_a_model_only_where_it_holds_geometry_the_viewer_reads() {
        let folder = std::env::temp_dir().join("rust-hover-preview-model-formats-dxf");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let sheet = folder.join("a-sheet.dxf");
        std::fs::write(
            &sheet,
            b"0\nSECTION\n2\nENTITIES\n0\nLWPOLYLINE\n90\n2\n0\nENDSEC\n0\nEOF\n",
        )
        .expect("a drawing");
        assert!(
            !claims_content(&sheet, None),
            "a sheet of 2D geometry is the render engine's, not a model"
        );

        let mesh = folder.join("a-mesh.dxf");
        std::fs::write(
            &mesh,
            b"0\nSECTION\n2\nENTITIES\n0\n3DFACE\n8\n0\n0\nENDSEC\n0\nEOF\n",
        )
        .expect("a drawing");
        assert!(
            claims_content(&mesh, None),
            "a drawing holding a face is the model the name promises"
        );

        // Verbatim DXF writes its group codes right-justified in three columns, so a drawing the
        // application saved opens `  0\r\nSECTION`: the padding is the file's own formatting, and
        // what is behind it is the same drawing.
        let padded_mesh = folder.join("a-padded-mesh.dxf");
        std::fs::write(
            &padded_mesh,
            b"  0\r\nSECTION\r\n  2\r\nENTITIES\r\n  0\r\n3DFACE\r\n  8\r\n0\r\n  0\r\nENDSEC\r\n  0\r\nEOF\r\n",
        )
        .expect("a drawing");
        assert!(
            claims_content(&padded_mesh, None),
            "a face written under a padded group code is the model it was unpadded"
        );

        let padded_sheet = folder.join("a-padded-sheet.dxf");
        std::fs::write(
            &padded_sheet,
            b"  0\r\nSECTION\r\n  2\r\nENTITIES\r\n  0\r\nLWPOLYLINE\r\n 90\r\n2\r\n  0\r\nENDSEC\r\n",
        )
        .expect("a drawing");
        assert!(
            !claims_content(&padded_sheet, None),
            "and a padded sheet is still the render engine's, as the unpadded one is"
        );

        // A BOM is not a group code: the front steps over one, so the body is read from the code
        // behind it and the drawing is answered as it would be without it.
        let bom_mesh = folder.join("a-bom-mesh.dxf");
        std::fs::write(
            &bom_mesh,
            "\u{feff}0\r\nSECTION\r\n2\r\nENTITIES\r\n0\r\n3DFACE\r\n8\r\n0\r\n0\r\nENDSEC\r\n",
        )
        .expect("a drawing");
        assert!(
            claims_content(&bom_mesh, None),
            "a face behind a BOM is the model the drawing holds"
        );

        let bom_sheet = folder.join("a-bom-sheet.dxf");
        std::fs::write(
            &bom_sheet,
            "\u{feff}0\r\nSECTION\r\n2\r\nENTITIES\r\n0\r\nLWPOLYLINE\r\n90\r\n2\r\n0\r\nENDSEC\r\n",
        )
        .expect("a drawing");
        assert!(
            !claims_content(&bom_sheet, None),
            "and a sheet behind one is the render engine's too"
        );

        // A pad is the few columns a code is written in, not an indent of any depth.
        assert!(
            !matcher("dxf")(b"        0\r\nSECTION", 0),
            "a front indented past the columns a code is written in is no drawing"
        );

        // The binary spelling has the front `is_dxf` names and no head the importer can read:
        // the claim is the front's first, and the body's after it.
        let binary = folder.join("a-binary.dxf");
        std::fs::write(
            &binary,
            b"AutoCAD Binary DXF\r\n\x1a\x00\x000\x00S\x00E\x00C\x00T",
        )
        .expect("a binary drawing");
        assert!(
            !claims_content(&binary, None),
            "a drawing written as binary is one the viewer cannot open"
        );

        // And a name that is not a drawing is still its own front's answer alone.
        let mesh_of_another_name = folder.join("mesh.stl");
        std::fs::write(&mesh_of_another_name, b"solid cube\nendsolid cube\n").expect("a mesh");
        assert!(claims_content(&mesh_of_another_name, None));

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// The rows and the built-in list are held to each other in both directions: a name added to
    /// the list without a matcher fails here, and so does a matcher left behind by a name that was
    /// taken out of it.
    #[test]
    fn the_table_says_a_matcher_for_every_name_the_list_carries_and_no_other() {
        let names = names();

        for name in &names {
            assert!(
                MODEL_FORMATS.iter().any(|(known, _)| known == name),
                "`{name}` is a name the model list carries and this table says nothing about it"
            );
        }
        for (name, _) in MODEL_FORMATS {
            assert!(
                names.contains(name),
                "`{name}` is not a name the model list carries, so nothing claims it"
            );
        }
        assert_eq!(
            MODEL_FORMATS.len(),
            names.len(),
            "a name written down twice is a matcher one of the two never reaches"
        );
    }
}
