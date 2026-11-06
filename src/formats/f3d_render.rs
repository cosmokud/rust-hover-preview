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
//! **The command line is this module's, and the window is not.** `launch_args` is the whole of
//! what f3d is told, and it is a pure function so that the one thing a wrong flag costs — an f3d
//! that exits with no window and a preview that silently shows nothing — is testable without a
//! desktop. What begins it, waits for its window and dresses it is `f3d_player`, beside the two
//! players this app already hosts.

use crate::config::config::TransparentBackground;
use crate::formats::codecs;
use once_cell::sync::Lazy;
use std::ffi::{OsStr, OsString};
use std::path::{Path, PathBuf};

/// The program inside an installation that draws a model. Also the image name a model's process
/// is recorded and verified under, beside `ffplay.exe` and `mpv.exe`.
pub const F3D_NAME: &str = "f3d.exe";

/// The box a model's hover is laid out at, in the shape a picture's preview takes: a scene has no
/// size of its own to be measured by — it is drawn at whatever box it is given, the way a page the
/// browser draws is — so what a hover asks for is this box scaled by the picture share (see
/// `scale_of_kind`).
pub const MODEL_WIDTH: u32 = 1280;
pub const MODEL_HEIGHT: u32 = 800;

/// The box a model's preview is laid out at (see `MODEL_WIDTH`).
pub fn model_box() -> (u32, u32) {
    (MODEL_WIDTH, MODEL_HEIGHT)
}

/// The backdrop a model is drawn over, as f3d names one: the app's own picture backdrop, which is
/// what stands behind every other kind of preview (see `routing::backdrop_of`).
///
/// **A checkerboard is answered with its darker square**, and a backdrop that shows the desktop
/// through with the same grey: f3d paints one colour and has no squares to paint, so the choice is
/// which of the app's own greys a scene reads best against, and that is the darker one.
pub fn background_argument(backdrop: TransparentBackground) -> String {
    // `0.565` is 144/255, the dark square of `checkerboard_color`.
    let value = match backdrop {
        TransparentBackground::White => "1.000,1.000,1.000",
        TransparentBackground::Black => "0.000,0.000,0.000",
        TransparentBackground::Checkerboard | TransparentBackground::Transparent => {
            "0.565,0.565,0.565"
        }
    };

    value.to_string()
}

/// What a launch tells f3d, as the command line it is begun with.
///
/// **The configuration is kept and the overlays are turned off one at a time.** `--no-config`
/// exists in f3d 4.0.1 and is not used: it drops the per-format configs as well as the stock one,
/// and those are what stand an `.stl` the right way up (`10_native.json` sets its `up` axis and
/// camera direction). Measured against the installed 4.0.1 — the same file rendered with
/// `--no-config` and with these switches is the same picture as the stock configuration's, with
/// the axis, the grid and the filename gone.
///
/// Every switch below was checked against `f3d --help` and by rendering with it: `-D` takes
/// `option=value` as one argument, an option f3d does not know is warned about and ignored, and
/// the names are the flat ones the config files themselves use (`filename`, not `ui.filename`).
/// The three `ui.` entries are the window's own furniture and take that prefix.
///
/// `--position` is the **decorated** window's origin and `--resolution` its client size, both
/// verified: the frame is stripped after the launch (see `f3d_player`), and where the window is
/// finally placed is this app's to say rather than the player's — which is why the origin is a
/// caller's argument and is **not the box**: the viewer shows its window itself, so it is begun off
/// every display and moved to the box once this app holds it (see `f3d_player::the_parking_origin`).
/// The size is the box's, always: the crop a host's thumbnail is cut from is that client (see
/// `f3d_player::note_the_box`).
///
/// **One effect is turned on, and it is the one the stock configuration leaves off.** A model that
/// carries no texture of its own — every `.stl`, most CAD exchanges — is one flat grey material
/// under a single stock light, which measured as a silhouette whose luminance varied by about
/// seven levels out of 255 across the whole body — a cut-out, not a solid (see
/// `the_launch_asks_for_depth_as_well_as_the_box`). `ambient-occlusion` is
/// the only switch that puts that back without touching a colour the file itself carries: it is
/// screen-space contact shading, so creases, cavities and the seams between parts darken, and a
/// plain grey mesh reads as a solid again. Tone mapping and anti-aliasing are deliberately *not*
/// repeated here — the stock configuration already sets both, and a flag that agrees with it is
/// noise that would have to be kept in step.
pub fn launch_args(
    path: &Path,
    box_: (i32, i32, i32, i32),
    position: (i32, i32),
    background: &str,
) -> Vec<OsString> {
    let (_, _, width, height) = box_;
    let (x, y) = position;
    let mut args: Vec<OsString> = vec![
        "--resolution".into(),
        format!("{},{}", width.max(1), height.max(1)).into(),
        "--position".into(),
        format!("{},{}", x, y).into(),
        "--multi-file-mode".into(),
        "single".into(),
    ];

    for option in [
        "ui.console=false",
        "ui.cheatsheet=false",
        "ui.drop_zone.enable=false",
        "axis=false",
        "grid=false",
        "filename=false",
        "hdri-filename=false",
        "notifications=false",
        "loading-progress=false",
        "animation-progress=none",
        // The one effect this app asks for rather than inherits (see `launch_args`).
        "ambient-occlusion=true",
    ] {
        args.push("-D".into());
        args.push(option.into());
    }

    args.push("-D".into());
    args.push(format!("background-color={background}").into());
    args.push(path.as_os_str().to_os_string());
    args
}

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

    /// **The launch is the file, the box, and every switch that keeps f3d's own furniture off the
    /// screen**, and nothing else is asserted: the flags themselves are the measurement (see
    /// `launch_args`), and what a test can hold is that the one that matters most — the stock
    /// configuration's axis, grid and filename — is turned off rather than the configuration being
    /// dropped for the per-format camera with it.
    #[test]
    fn the_launch_asks_for_the_box_and_turns_every_overlay_off() {
        let args = launch_args(
            Path::new("a-model.stl"),
            (10, 20, 640, 480),
            (10, 20),
            "0.5,0.5,0.5",
        );
        let asked: Vec<String> = args
            .iter()
            .map(|arg| arg.to_string_lossy().into_owned())
            .collect();
        let has = |argument: &str| asked.iter().any(|asked| asked == argument);

        assert!(has("640,480"), "the client size asked for: {asked:?}");
        assert!(has("10,20"), "and where the window is asked to be");
        assert!(has("single"), "one file is one scene");
        for overlay in ["axis=false", "grid=false", "filename=false"] {
            assert!(has(overlay), "{overlay} is turned off: {asked:?}");
        }
        assert!(has("background-color=0.5,0.5,0.5"), "the app's backdrop");
        assert!(
            !has("--no-config"),
            "the per-format camera configs are kept (see `launch_args`)"
        );
        assert_eq!(
            asked.last().map(String::as_str),
            Some("a-model.stl"),
            "the file is last, which is where f3d takes it"
        );
    }

    /// **The size is the box's and the place is the caller's**, which are two different numbers on the
    /// parked launch a hover or a pin is actually begun with (see `f3d_player::the_parking_origin`): a
    /// viewer's window is begun off every display and moved to the box once this app holds it, so a flag
    /// read the other way round would put the decorated window back on the screen for a composition
    /// frame — the flash this road exists to not have.
    #[test]
    fn the_parked_launch_takes_the_size_from_the_box_and_the_place_from_the_caller() {
        let asked: Vec<String> = launch_args(
            Path::new("a-model.stl"),
            (10, 20, 640, 480),
            (12_345, -400),
            "0",
        )
        .iter()
        .map(|arg| arg.to_string_lossy().into_owned())
        .collect();

        assert!(
            asked.iter().any(|arg| arg == "640,480"),
            "the box's client size is what the viewer renders at: {asked:?}"
        );
        assert!(
            asked.iter().any(|arg| arg == "12345,-400"),
            "and the place is the parking origin it was handed: {asked:?}"
        );
        assert!(
            !asked.iter().any(|arg| arg == "10,20"),
            "the box's own origin is never asked for, which is what keeps f3d's window off the \
             screen: {asked:?}"
        );
    }

    /// **Depth is the one effect the launch adds**, and it is asked for by the name f3d's own
    /// configuration files use: a plain grey mesh is lit as a near-uniform silhouette without it,
    /// and the shading it buys is the only thing that makes one read as a solid rather than a cut
    /// out. Nothing else is asked for that the stock configuration already sets — a switch that
    /// agreed with it could drift out of step with the next f3d and would never be noticed.
    #[test]
    fn the_launch_asks_for_depth_as_well_as_the_box() {
        let asked: Vec<String> =
            launch_args(Path::new("a-model.stl"), (0, 0, 640, 480), (0, 0), "0")
                .iter()
                .map(|arg| arg.to_string_lossy().into_owned())
                .collect();

        assert!(
            asked.iter().any(|arg| arg == "ambient-occlusion=true"),
            "contact shading is switched on: {asked:?}"
        );
        for inherited in ["tone-mapping=", "anti-aliasing=", "light-intensity="] {
            assert!(
                !asked.iter().any(|arg| arg.starts_with(inherited)),
                "{inherited} is the stock configuration's to set: {asked:?}"
            );
        }
    }

    /// A model's backdrop is the app's own, colour for colour: the two flat ones as themselves,
    /// and the two a viewer cannot paint as the darker of the squares a picture is drawn over.
    #[test]
    fn the_backdrop_is_the_apps_own_colour() {
        assert_eq!(
            background_argument(TransparentBackground::White),
            "1.000,1.000,1.000"
        );
        assert_eq!(
            background_argument(TransparentBackground::Black),
            "0.000,0.000,0.000"
        );
        assert_eq!(
            background_argument(TransparentBackground::Checkerboard),
            "0.565,0.565,0.565"
        );
        assert_eq!(
            background_argument(TransparentBackground::Transparent),
            background_argument(TransparentBackground::Checkerboard),
            "a viewer has no squares and no desktop to show through"
        );
    }

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
