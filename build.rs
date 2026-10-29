use std::env;
use std::process::{Command, Stdio};

fn main() {
    // Avoid Cargo scanning the entire repo tree (can fail on locked temp dirs).
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=Cargo.toml");
    println!("cargo:rerun-if-changed=Cargo.lock");
    println!("cargo:rerun-if-changed=src");
    println!("cargo:rerun-if-changed=assets/icon.ico");

    // One target dir serves every checkout, and cargo's fingerprint for this package
    // cannot tell two of them apart. The checkout's own path is in the environment
    // (written by .githooks/post-checkout), so a build in another checkout changes
    // this value and recompiles, instead of reusing a binary built from other sources.
    println!("cargo:rerun-if-env-changed=RHP_WORKTREE");

    // A release build links over `rust-hover-preview.exe`, which Windows will
    // not allow while the app is up: the link fails with "Access is denied"
    // and the build has to be run again. Ending a running copy here, before
    // the crate is compiled and linked, is what lets `cargo build --release`
    // replace the binary in one pass. Watching `src` above is what brings this
    // script back for the rebuilds that need it.
    #[cfg(target_os = "windows")]
    if env::var("PROFILE").as_deref() == Ok("release") {
        stop_running_app();
    }

    // The resources belong to the *target*, not the host that runs this script. `cfg!(windows)`
    // here answers for the host, and a build script compiled on Linux would then skip the icon,
    // the version block and the manifest — silently, because nothing else about the link fails.
    // The target is therefore asked for by name, so that a cross build from another system
    // embeds exactly what a build on Windows does.
    if env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows") {
        let mut res = winres::WindowsResource::new();
        // Away from Windows there is no SDK laid out where winres looks for one, so the
        // directory holding an `rc.exe`-compatible compiler comes from the machine instead
        // of the tree. It is never committed for the same reason the shared target dir is
        // not: it names a path on the machine doing the build, not on the one reading this.
        if let Ok(toolkit) = env::var("RHP_RC_TOOLKIT") {
            res.set_toolkit_path(&toolkit);
        }
        res.set_icon("assets/icon.ico");
        res.set("ProductName", "Rust Hover Preview");
        res.set("FileDescription", "Rust Hover Preview");
        res.set("LegalCopyright", "Copyright 2026");
        // And the version lie with it. An executable with no manifest is given the compatibility
        // of the oldest Windows there is: `GetVersionExW` answers 6.2 on a machine far past it and
        // the pre-Win8 semantics come with it, and one of those is that a `WS_CHILD` window may
        // not be `WS_EX_LAYERED` — `CreateWindowExW` refuses the pair with error 87, which is what
        // buried a pinned mpv film's chrome (the film's host is that pair; see `pin_film_host`).
        // Declaring the Windows this app supports is the whole of what turns those semantics off,
        // and it is asked for as raw XML rather than through a convenience form, so that nothing
        // rides along with it — no common-controls dependency, no DPI declaration, no requested
        // execution level — because every one of those would change how this app is treated in a
        // way nothing here has measured.
        res.set_manifest(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
    <application>
      <supportedOS Id="{4a2f28c8-2506-4483-a6d7-18b95a9b46ea}"/>
      <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
    </application>
  </compatibility>
</assembly>
"#,
        );

        // A failure here is not a warning to read later. The icon and the version block are what
        // the Shell draws this app by and what its own tray loads at run time (see
        // `shell::app_icon`), so a binary built without them is a binary that wears the system's
        // placeholder — and nothing else about it is wrong, so nothing else says a word.
        if let Err(e) = res.compile() {
            panic!("could not compile the Windows resources: {e}");
        }

        // And what the compiler produced has to be handed to the *binary's* link. Cargo passes a
        // build script's `rustc-link-lib` to the package's library target rather than to the
        // binary that links it, so a `src/lib.rs` appearing beside `main.rs` is enough to drop
        // both the icon and the version block out of the executable with every exit code still
        // zero (`tests/executable_icon` is what asks, and there is no other symptom until a
        // build is run). Named here as the binary's own link argument, the file arrives whatever
        // else the package grows, and winres's own `rustc-link-lib` line stays where it is: the
        // same file, named for the library target.
        let resource_file = std::path::Path::new(&env::var("OUT_DIR").expect("cargo sets OUT_DIR"))
            .join("resource.lib");
        // The file is an archive, and a linker pulls a member out of one only to satisfy a
        // symbol it is missing — a resource object defines no symbols at all, so off Windows
        // the member is never pulled and the icon, the version block and the manifest would be
        // dropped with every exit code still zero. `/WHOLEARCHIVE` is what takes the member
        // whether it is wanted or not. Windows needs no such argument: the RC compiler there
        // writes the same archive in a form its linker takes regardless (see `tests/executable_icon`).
        let whole_archive = if cfg!(windows) { "" } else { "/WHOLEARCHIVE:" };
        println!(
            "cargo:rustc-link-arg-bin={}={}{}",
            env::var("CARGO_PKG_NAME").expect("cargo sets CARGO_PKG_NAME"),
            whole_archive,
            resource_file.display()
        );
    }
}

/// Ends a running copy of the app so a release build can replace its binary.
/// Nothing is reported when the app is not running, which is the usual case.
#[cfg(target_os = "windows")]
fn stop_running_app() {
    use std::os::windows::process::CommandExt;

    /// CREATE_NO_WINDOW, so a build started from an IDE shows no console flash.
    const CREATE_NO_WINDOW: u32 = 0x0800_0000;

    let image = concat!(env!("CARGO_PKG_NAME"), ".exe");
    let stopped = Command::new("taskkill")
        .args(["/F", "/IM", image])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(CREATE_NO_WINDOW)
        .status()
        .is_ok_and(|status| status.success());

    if stopped {
        println!("cargo:warning=Stopped the running {image} before rebuilding");
    }
}
