use std::env;
use std::process::{Command, Stdio};

fn main() {
    // Avoid Cargo scanning the entire repo tree (can fail on locked temp dirs).
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=Cargo.toml");
    println!("cargo:rerun-if-changed=Cargo.lock");
    println!("cargo:rerun-if-changed=src");
    println!("cargo:rerun-if-changed=assets/icon.ico");

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

    // Only run on Windows
    #[cfg(target_os = "windows")]
    {
        let mut res = winres::WindowsResource::new();
        res.set_icon("assets/icon.ico");
        res.set("ProductName", "Rust Hover Preview");
        res.set("FileDescription", "Rust Hover Preview");
        res.set("LegalCopyright", "Copyright 2026");
        if let Err(e) = res.compile() {
            eprintln!("Warning: Failed to compile Windows resources: {}", e);
        }
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
