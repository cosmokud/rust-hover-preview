//! Design documents rendered by an installed LibreOffice.
//!
//! A CorelDRAW document is a vector drawing, and what this app reads of one by itself is
//! the picture the application saved for a file manager — small, and soft when the
//! preview is enlarged. The drawing itself is in a proprietary object format three
//! generations deep, which is why this app has no reader for it. LibreOffice has one: its
//! import filters ship the Document Liberation Project's libraries, `libcdr` among them,
//! and those parse the drawing rather than its thumbnail. So where LibreOffice is
//! installed it is asked first, and what it hands back is a PDF — a page of vector data,
//! which the PDF path already draws at whatever size a preview is shown at. Where it is
//! not installed, nothing changes: the readers of the pictures the files carry are the
//! fallback, and they are also what answers a document the engine cannot read.
//!
//! Nothing is bundled with this app and nothing is linked against: the engine is the
//! user's own installation, looked for where it installs — and beside `config.ini` for a
//! portable copy — and run as the user runs it. What that costs is a launch, which is
//! seconds, so a conversion happens once per document: the PDF it wrote is kept under
//! [`AppConfig::rendered_dir`], named for the document's path and the version of it that
//! was converted, and every hover after the first is a read of that file. One conversion
//! runs at a time, because one LibreOffice at a time is what its own profile allows.
//!
//! The engine is asked about the names its filters read and no others — the formats of
//! the libraries above, CorelDRAW's and the rest — so a Photoshop document, a Krita
//! project or a Procreate file never pays for a launch that could not answer.

use crate::config::AppConfig;
use once_cell::sync::Lazy;
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::Mutex;
use std::time::{Duration, Instant};

/// How long a conversion is given before it is ended. A conversion of the documents this
/// is for takes seconds; the first one after an install also writes the engine's own
/// profile, which is why the wait is generous rather than short.
const CONVERSION_TIMEOUT: Duration = Duration::from_secs(60);
/// How often the wait above looks.
const CONVERSION_POLL: Duration = Duration::from_millis(100);
/// What a name is remembered as when the engine would not draw it: the conversion is not
/// tried again for that version of the document, because a name this app was wrong about —
/// one in the list the engine has no filter for — would otherwise start an engine on every
/// hover to reach the same answer.
const REFUSED_SUFFIX: &str = "none";

/// Where LibreOffice keeps its program, for the two places it installs and for a portable
/// copy a user may have put beside `config.ini`.
fn soffice() -> Option<PathBuf> {
    static FOUND: Lazy<Option<PathBuf>> = Lazy::new(|| {
        for candidate in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ] {
            let path = Path::new(candidate);
            if path.is_file() {
                return Some(path.to_path_buf());
            }
        }

        let portable = AppConfig::rendered_dir()?
            .parent()?
            .join("libreoffice")
            .join("program")
            .join("soffice.exe");

        portable.is_file().then_some(portable)
    });

    FOUND.clone()
}

/// Whether an engine is installed to render these documents with.
pub fn available() -> bool {
    soffice().is_some()
}

/// The rendered page of `path`, converting the document if it has not been converted
/// before: a PDF under the app's own folder, which the PDF path reads the way it reads
/// any other.
///
/// `None` is the answer for a name the configured list does not hold, for a machine without
/// the engine, and for a document it could not convert — each of which leaves the preview to
/// the readers of the picture the file carries.
pub fn pdf_for(path: &Path) -> Option<PathBuf> {
    imports(path).then(|| rendered(path)).flatten()
}

/// The same for a document the `[libre]` list does not hold — one of the Office kind, asked
/// for here only where the Office engine is not installed. The caller decides that: what a
/// name means is the lists' business, and this is the engine that draws whatever it is
/// given.
pub fn pdf_for_office(path: &Path) -> Option<PathBuf> {
    rendered(path)
}

/// Whether the list this app keeps says the engine is the one to draw `path`.
fn imports(path: &Path) -> bool {
    crate::libre_formats::is_libre_file(path)
}

/// The rendered page of a document, converting it if it has not been converted before.
fn rendered(path: &Path) -> Option<PathBuf> {
    let program = soffice()?;
    let page = rendered_path(path)?;
    if usable(&page) {
        return Some(page);
    }
    if refused(&page).is_some() {
        return None;
    }

    // One engine at a time, and the lock is held across the conversion: LibreOffice's
    // profile is a single seat, so a second run beside the first would wait on it anyway.
    let _converting = CONVERTING.lock().ok()?;
    if usable(&page) {
        return Some(page);
    }
    if refused(&page).is_some() {
        return None;
    }

    if convert(&program, path, &page).is_none() {
        // An engine that would not draw this document is not asked again: what it answered
        // is written down beside the page it did not write.
        std::fs::write(refused_path(&page), b"").ok();
        return None;
    }

    Some(page)
}

/// Where the rendered page of `path` is kept: named for the document — its path, the
/// version of it that was converted, and nothing else — so a document saved again is
/// rendered again and a document that has not been is a read.
fn rendered_path(path: &Path) -> Option<PathBuf> {
    let folder = AppConfig::rendered_dir()?;
    std::fs::create_dir_all(&folder).ok()?;

    let mut hasher = DefaultHasher::new();
    path.to_string_lossy().to_lowercase().hash(&mut hasher);
    let metadata = std::fs::metadata(path).ok();
    metadata.as_ref().map(|metadata| metadata.len()).hash(&mut hasher);
    metadata
        .and_then(|metadata| metadata.modified().ok())
        .and_then(|modified| modified.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|since| since.as_secs())
        .hash(&mut hasher);

    Some(folder.join(format!("{:016x}.pdf", hasher.finish())))
}

/// Whether a rendered page is there and is a PDF: what is read back is a file this app
/// wrote, so the header is the check and not a re-render.
fn usable(rendered: &Path) -> bool {
    let Ok(mut file) = std::fs::File::open(rendered) else {
        return false;
    };
    let mut header = [0u8; 5];
    std::io::Read::read_exact(&mut file, &mut header).is_ok() && &header == b"%PDF-"
}

/// The mark left beside a page that was not written, for the same document and the same
/// version of it, or nothing when the engine has not been asked about it yet.
fn refused(page: &Path) -> Option<PathBuf> {
    let refused = refused_path(page);

    refused.is_file().then_some(refused)
}

fn refused_path(page: &Path) -> PathBuf {
    page.with_extension(REFUSED_SUFFIX)
}

/// Convert `source` into `rendered`, by running the engine the way a user would: headless,
/// with a profile of this app's own so that a LibreOffice the user has open is untouched,
/// and with a wait that ends rather than holding a hover for good.
fn convert(program: &Path, source: &Path, rendered: &Path) -> Option<()> {
    let folder = rendered.parent()?;
    let stage = folder.join("stage");
    // Whatever a run before this one left behind is not read: the engine names what it
    // writes after what it was given, and only the file of this conversion is looked for.
    std::fs::remove_dir_all(&stage).ok();
    std::fs::create_dir_all(&stage).ok()?;

    let profile = folder.join("profile");
    let profile_url = format!("file:///{}", profile.to_string_lossy().replace('\\', "/"));

    let mut child = Command::new(program)
        .arg("--headless")
        .arg("--norestore")
        .arg(format!("-env:UserInstallation={profile_url}"))
        .arg("--convert-to")
        .arg("pdf:draw_pdf_Export")
        .arg("--outdir")
        .arg(&stage)
        .arg(source)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    // A process this app started is put in the job every other engine is put in, so one
    // that outlives the app does not outlive it by much.
    crate::engine_processes::adopt(child.id());

    if !wait(&mut child, CONVERSION_TIMEOUT) {
        return None;
    }

    // The engine names what it wrote after what it was given, in the folder it was told —
    // and the name is the document's own, dots and all, which is why it is appended to
    // rather than replaced: a file called `drawing.v2.cdr` is written as
    // `drawing.v2.pdf`, not as `drawing.pdf`.
    let stem = source.file_stem()?;
    let written = stage.join(format!("{}.pdf", stem.to_string_lossy()));
    let page = std::fs::read(&written).ok()?;
    if !page.starts_with(b"%PDF-") {
        return None;
    }

    std::fs::write(rendered, &page).ok()?;
    std::fs::remove_file(&written).ok();
    prune(folder);

    Some(())
}

/// Wait for a process, ending it rather than waiting past `limit`.
fn wait(child: &mut Child, limit: Duration) -> bool {
    let deadline = Instant::now() + limit;
    loop {
        match child.try_wait() {
            Ok(Some(_)) => return true,
            Ok(None) if Instant::now() < deadline => std::thread::sleep(CONVERSION_POLL),
            _ => {
                let _ = child.kill();
                let _ = child.wait();
                return false;
            }
        }
    }
}

/// Keep the folder within the size the `Performance → Cache → Libre` setting names, oldest
/// first: a rendered page is built again from its document whenever it is wanted, so nothing
/// here is worth growing a folder for. A budget of nothing drops every page, which is what
/// the setting means — nothing is kept between hovers.
fn prune(folder: &Path) {
    let budget = crate::CONFIG
        .lock()
        .map(|config| config.libre_cache_mb as u64 * 1024 * 1024)
        .unwrap_or(0);
    let Ok(entries) = std::fs::read_dir(folder) else {
        return;
    };

    let mut pages: Vec<(std::time::SystemTime, u64, PathBuf)> = entries
        .flatten()
        .filter(|entry| {
            entry
                .path()
                .extension()
                .is_some_and(|extension| extension == "pdf" || extension == REFUSED_SUFFIX)
        })
        .filter_map(|entry| {
            let metadata = entry.metadata().ok()?;
            Some((metadata.modified().ok()?, metadata.len(), entry.path()))
        })
        .collect();

    let mut total: u64 = pages.iter().map(|(_, size, _)| size).sum();
    if total <= budget {
        return;
    }

    pages.sort_by_key(|(modified, _, _)| *modified);
    for (_, size, path) in pages {
        if total <= budget {
            break;
        }
        if std::fs::remove_file(&path).is_ok() {
            total = total.saturating_sub(size);
        }
    }
}

/// The one conversion running at a time, for the reason the engine has one profile.
static CONVERTING: Lazy<Mutex<()>> = Lazy::new(|| Mutex::new(()));

#[cfg(test)]
mod tests {
    use super::*;

    /// The engine is never started for a name its own filters do not read. A Photoshop
    /// document, a Krita project, a Sketch file and a Procreate document are this app's
    /// own readers' business, and asking an office suite about one would cost a launch and
    /// answer nothing.
    #[test]
    fn asks_the_engine_only_about_the_names_it_reads() {
        for name in ["poster.psd", "painting.kra", "design.sketch", "art.procreate", "icon.svg"]
        {
            assert_eq!(pdf_for(Path::new(name)), None, "`{name}` is not one of its formats");
        }
    }

    /// And the names it does read are the CorelDRAW family and the formats of the same
    /// libraries, whatever case they are written in.
    #[test]
    fn reads_coreldraw_and_the_formats_beside_it() {
        for name in ["logo.cdr", "drawing.CDR", "artwork.cmx", "poster.pub", "plan.vsd"] {
            assert!(imports(Path::new(name)), "`{name}` is one of its formats");
        }
    }
}
