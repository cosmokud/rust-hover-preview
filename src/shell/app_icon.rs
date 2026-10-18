//! The icon this app wears wherever the Shell draws it.
//!
//! It is the icon `build.rs` compiles into the executable rather than one drawn here, because an
//! icon is a file the Shell reads on its own: a taskbar button, a tray icon and a window's
//! Alt+Tab frame are all drawn by the Shell from the resource and never from anything this app
//! paints. The resource ID is the one the build script writes it under.

use windows::core::PCWSTR;
use windows::Win32::Foundation::HINSTANCE;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    IMAGE_ICON, LR_DEFAULTSIZE, LR_SHARED, LoadImageW, HICON,
};

/// The application's own icon, or the system's application icon where the resource cannot be
/// loaded.
///
/// `LR_DEFAULTSIZE` asks for the size the Shell draws an icon at rather than for a pixel size
/// fixed here, and `LR_SHARED` borrows the module's own copy instead of handing out one that
/// would have to be destroyed — which is what makes this callable from everywhere it is wanted
/// (a window class asks for it as it registers, and a class is registered once) without leaving
/// an icon to answer for.
///
/// The fallback is the ordinal the Shell keeps an application's default icon under. It is not
/// this app's icon, so it is only ever the answer to a broken build; the tray already draws
/// attention to nothing when it is used, and a taskbar button wearing it is the same admission.
pub(crate) unsafe fn app_icon() -> HICON {
    let Ok(hmodule) = GetModuleHandleW(None) else {
        return HICON::default();
    };

    match LoadImageW(
        HINSTANCE(hmodule.0),
        // `MAKEINTRESOURCEW(1)`: with the high word zero, `LoadImageW` reads this value as the ID
        // of the resource to load rather than as the address of a name, so it is an ID here and
        // not a pointer.
        PCWSTR(std::ptr::without_provenance(1)),
        IMAGE_ICON,
        0,
        0,
        LR_DEFAULTSIZE | LR_SHARED,
    ) {
        Ok(icon) if !icon.0.is_null() => HICON(icon.0),
        _ => {
            // Safety: asks for the Shell's own application icon. The ordinal is the one
            // `IDI_APPLICATION` is spelled by, since the constant is not a pointer here either.
            match LoadImageW(
                None,
                PCWSTR(32512 as *const u16),
                IMAGE_ICON,
                0,
                0,
                LR_DEFAULTSIZE | LR_SHARED,
            ) {
                Ok(icon) => HICON(icon.0),
                Err(_) => HICON::default(),
            }
        }
    }
}
