//! The icon this app wears is read out of the executable itself. The Shell draws a taskbar
//! button and a taskbar preview's header from it, and the tray icon is that same resource loaded
//! by this app at run time (see `shell::app_icon`), so a binary that was linked without it is not
//! a build that is missing a picture — it is an app that shows the system's placeholder in the
//! tray and no icon at all in a preview's header, while its taskbar button wears something
//! because the Shell draws a button even for a file with no icon in it.
//!
//! The resources are compiled by `build.rs` and handed to the link as a separate file, which is
//! the one step of a build that can go missing while every compiler exit code stays zero: no
//! source is wrong, so nothing fails, and only the executable is different. This asks the
//! executable as cargo built it, which is where the Shell will look.

use windows::core::{HSTRING, PCWSTR};
use windows::Win32::Foundation::{FreeLibrary, HANDLE, HMODULE};
use windows::Win32::System::LibraryLoader::{
    FindResourceW, LoadLibraryExW, LoadResource, LockResource, SizeofResource,
    LOAD_LIBRARY_AS_DATAFILE,
};
use windows::Win32::UI::WindowsAndMessaging::{RT_GROUP_ICON, RT_MANIFEST, RT_VERSION};

/// The ID `build.rs` compiles both resources under — `1 ICON` and `1 VERSIONINFO` in the resource
/// script winres writes — which is also the ID the app loads its own icon by at run time (see
/// `shell::app_icon`).
const ITS_OWN: usize = 1;

/// The keep of `MAKEINTRESOURCEW`: with the high word zero, the resource calls below read this
/// value as the ID of a resource rather than as the address of a name.
fn ordinal(id: usize) -> PCWSTR {
    PCWSTR(id as *const u16)
}

/// Whether the executable at `path` carries a resource of `kind` under `id`.
///
/// It is loaded as a data file rather than as a module: what is being read is a file's resources,
/// and nothing here wants that file's code in this process.
fn carries(path: &str, id: usize, kind: PCWSTR) -> bool {
    unsafe {
        let image: HMODULE = LoadLibraryExW(
            &HSTRING::from(path),
            HANDLE(std::ptr::null_mut()),
            LOAD_LIBRARY_AS_DATAFILE,
        )
        .expect("the executable cargo built");
        let found = FindResourceW(image, ordinal(id), kind);
        let _ = FreeLibrary(image);
        !found.0.is_null()
    }
}

#[test]
fn the_executable_carries_its_own_icon_and_its_own_version_block() {
    let the_exe = env!("CARGO_BIN_EXE_rust-hover-preview");

    assert!(
        carries(the_exe, ITS_OWN, RT_GROUP_ICON),
        "the executable has no icon in it, so the tray and a preview's header have none to draw \
         (the resources compiled by `build.rs` did not reach the link: see the link argument it \
         prints, and any `src/lib.rs` that would take it away from the binary)"
    );
    assert!(
        carries(the_exe, ITS_OWN, RT_VERSION),
        "the executable has no version block in it: properties and the task manager would show it \
         as a file with no name"
    );
}

/// The supportedOS GUIDs the manifest declares, one for Windows 8, one for 8.1 and one for
/// Windows 10 and 11. Their absence is not a missing picture: an executable with no manifest is
/// given the compatibility of the oldest Windows there is, the version lie comes with it, and one
/// of its pre-Win8 semantics refuses a `WS_CHILD` window made `WS_EX_LAYERED` — the pair a
/// pinned mpv film's picture is drawn in (see `pin_film_host`).
const SUPPORTED_OS: [&str; 3] = [
    "4a2f28c8-2506-4483-a6d7-18b95a9b46ea",
    "1f676c76-80e1-4239-95bb-83d0f6d0da78",
    "8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a",
];

/// The bytes of the executable's own manifest, or nothing where it has none. Read as a data file,
/// like `carries`, and answered as text because what a manifest says is text.
fn the_manifest(path: &str) -> Option<String> {
    unsafe {
        let image: HMODULE = LoadLibraryExW(
            &HSTRING::from(path),
            HANDLE(std::ptr::null_mut()),
            LOAD_LIBRARY_AS_DATAFILE,
        )
        .expect("the executable cargo built");
        let found = FindResourceW(image, ordinal(ITS_OWN), RT_MANIFEST);
        let bytes = (!found.is_invalid())
            .then_some(found)
            .and_then(|found| LoadResource(image, found).ok())
            .and_then(|held| {
                let address = LockResource(held);
                let size = SizeofResource(image, found);
                (size > 0 && !address.is_null()).then(|| {
                    std::slice::from_raw_parts(address.cast::<u8>(), size as usize).to_vec()
                })
            })
            .map(|bytes| String::from_utf8_lossy(&bytes).into_owned());
        let _ = FreeLibrary(image);
        bytes
    }
}

#[test]
fn the_executable_declares_the_windows_it_supports() {
    let the_exe = env!("CARGO_BIN_EXE_rust-hover-preview");

    let manifest = the_manifest(the_exe).expect("the executable has no manifest in it");
    for guid in SUPPORTED_OS {
        assert!(
            manifest.contains(guid),
            "the manifest does not declare {guid}, so the process is version-lied into \
             pre-Win8 semantics"
        );
    }
}
