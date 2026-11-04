//! Three-dimensional models produced by an installed f3d.
//!
//! A model file — a mesh, a CAD exchange document, a point cloud, a scientific dataset — is a
//! scene rather than a picture: nothing about it can be decoded into a frame without a renderer,
//! and this app carries none. f3d is the renderer that opens all of them, and what it hands back
//! is a picture of the model, which is the one thing this app already knows how to draw.
//!
//! What is asked of it is the program the installer writes, `f3d.exe`, run the way a user runs it.
//! Every name it is asked about is one its own reader list declares — see `lists::F3D` and
//! `DEFAULT_F3D_EXTENSIONS`, which is that list minus the names another kind of this app has
//! already claimed — so a picture or a document never pays for a launch that could not answer.
//!
//! **Nothing is launched here yet.** The picture f3d draws is shown in a window of its own, which
//! this module is the detection half of: what a machine has, and which of its installations is the
//! one to run. The host window, the mouse the model turns under and the process a hover starts are
//! the phases after this one.

use crate::formats::codecs;
use once_cell::sync::Lazy;
use std::ffi::OsStr;
use std::path::{Path, PathBuf};

/// The program inside an installation that draws a model.
const F3D_NAME: &str = "f3d.exe";

/// The folder the Windows installer writes, beside the other installations a machine may have.
///
/// Joined onto whatever `%ProgramFiles%` holds rather than written as a literal drive, the way
/// every other install folder of this app's is.
const INSTALL_FOLDER: &str = r"f3d\bin";

/// Whether an installation is here to draw models with.
pub fn available() -> bool {
    program().is_some()
}

/// Whether f3d has already turned `path` down.
///
/// It never has, and the answer is a constant for a reason worth stating: f3d 4.0.1 was measured
/// against every extension it lists, and it opened every one of them. What that makes the list is
/// the whole of the question — a name outside it is never started at all — so a file that reaches
/// this cannot be one the engine refuses.
pub fn refused(_path: &Path) -> bool {
    false
}

/// The `f3d.exe` a launch runs: the path the setting names while it is still a file, and the
/// first installation this machine has otherwise.
///
/// The choice is a path the user set and the installations are only this app's defaults, so a
/// chosen program that is not one of them — a portable copy, a build of their own — is still the
/// one to run (see `resolve`).
pub fn program() -> Option<PathBuf> {
    static FOUND: Lazy<Option<PathBuf>> = Lazy::new(|| {
        let chosen = crate::CONFIG
            .lock()
            .ok()
            .map(|config| PathBuf::from(&config.f3d_path));

        resolve(
            chosen.as_deref(),
            // The two places `Command::new` looks in, which f3d is found by on most machines.
            codecs::find_program(F3D_NAME),
            install_folder(std::env::var_os("ProgramFiles").as_deref())
                .filter(|program| program.is_file()),
        )
    });

    FOUND.clone()
}

/// Which of the three answers to run, in the order they are preferred: the chosen path while it is
/// still a file, then the one a `PATH` lookup found, then the folder the installer wrote.
///
/// A chosen path that has gone is not fatal: what is installed is run instead, and the setting is
/// left as the user wrote it for a machine they put it back on.
fn resolve(
    chosen: Option<&Path>,
    on_path: Option<PathBuf>,
    installed: Option<PathBuf>,
) -> Option<PathBuf> {
    chosen
        .filter(|chosen| chosen.is_file())
        .map(Path::to_path_buf)
        .or(on_path)
        .or(installed)
}

/// The program the Windows installer writes: `%ProgramFiles%\f3d\bin\f3d.exe`.
fn install_folder(program_files: Option<&OsStr>) -> Option<PathBuf> {
    Some(
        Path::new(program_files?)
            .join(INSTALL_FOLDER)
            .join(F3D_NAME),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The three places are preferred in the order a launch depends on: a path the user set is
    /// their choice and wins wherever it points, `f3d.exe` on the `PATH` comes next, and the
    /// installer's own folder is the fallback.
    ///
    /// Every one of the three is handed in rather than read, so nothing here touches the `PATH` or
    /// a real `Program Files`: the machine this runs on has an f3d of its own, and what a test
    /// answers about a machine must not be what that machine happens to hold.
    #[test]
    fn the_chosen_program_is_the_one_run_and_the_rest_are_what_it_falls_back_to() {
        let chosen = std::env::current_exe().expect("this test's own executable");
        let gone = std::env::temp_dir().join(format!(
            "rust-hover-preview-f3d-that-is-gone-{}",
            std::process::id()
        ));
        let on_path = PathBuf::from("a-path-folder").join(F3D_NAME);
        let installed = PathBuf::from("an-install").join(F3D_NAME);

        assert!(!gone.is_file(), "the chosen path is one that is not there");

        assert_eq!(
            resolve(
                Some(&chosen),
                Some(on_path.clone()),
                Some(installed.clone())
            ),
            Some(chosen.clone()),
            "a chosen path that is still a file is the one to run"
        );
        assert_eq!(
            resolve(Some(&gone), Some(on_path.clone()), Some(installed.clone())),
            Some(on_path.clone()),
            "a chosen path that has gone falls back to the PATH"
        );
        assert_eq!(
            resolve(Some(&gone), None, Some(installed.clone())),
            Some(installed.clone()),
            "and to the installer's folder where the PATH has none"
        );

        // Which is the whole of what an availability question answers: nothing chosen, nothing
        // found and nothing installed is no engine, and any one of them is one.
        assert_eq!(
            resolve(None, None, None),
            None,
            "a machine with no f3d anywhere has no model preview"
        );
        assert_eq!(
            resolve(Some(&chosen), None, None),
            Some(chosen),
            "a machine whose f3d is the path the user set has one"
        );
        assert!(
            resolve(None, Some(on_path), None).is_some(),
            "and one found on the PATH is one too"
        );
    }

    /// The installer's folder follows whichever root the environment gives, and a machine that
    /// names no `Program Files` at all has no such folder to look in.
    #[test]
    fn the_install_folder_follows_the_root_it_is_given() {
        assert_eq!(
            install_folder(Some(OsStr::new("a-root"))),
            Some(PathBuf::from("a-root").join(INSTALL_FOLDER).join(F3D_NAME)),
            "the install's own folder is joined onto the environment's root"
        );
        assert_eq!(install_folder(None), None);
    }
}
