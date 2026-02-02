fn main() {
    // Only run on Windows
    #[cfg(target_os = "windows")]
    {
        let mut res = winres::WindowsResource::new();
        res.set_icon("assets/icon.ico");
        res.set("ProductName", "Rust Hover Preview");
        res.set("FileDescription", "Image and video preview on hover in Windows Explorer");
        res.set("LegalCopyright", "Copyright 2026");
        if let Err(e) = res.compile() {
            eprintln!("Warning: Failed to compile Windows resources: {}", e);
        }
    }
}
