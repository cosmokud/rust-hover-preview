use std::env;
use windows::core::{w, PCWSTR};
use windows::Win32::System::Registry::{
    RegCloseKey, RegDeleteValueW, RegOpenKeyExW, RegSetValueExW, HKEY, HKEY_CURRENT_USER,
    KEY_SET_VALUE, REG_SZ,
};

const STARTUP_KEY: PCWSTR = w!(r"Software\Microsoft\Windows\CurrentVersion\Run");
const APP_NAME: PCWSTR = w!("RustHoverPreview");

pub fn enable_startup() {
    unsafe {
        let mut hkey: HKEY = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, STARTUP_KEY, 0, KEY_SET_VALUE, &mut hkey).is_ok() {
            if let Ok(exe_path) = env::current_exe() {
                let exe_path_wide: Vec<u16> = exe_path
                    .to_string_lossy()
                    .encode_utf16()
                    .chain(std::iter::once(0))
                    .collect();

                let _ = RegSetValueExW(
                    hkey,
                    APP_NAME,
                    0,
                    REG_SZ,
                    Some(&exe_path_wide.align_to::<u8>().1),
                );
            }
            let _ = RegCloseKey(hkey);
        }
    }
}

pub fn disable_startup() {
    unsafe {
        let mut hkey: HKEY = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, STARTUP_KEY, 0, KEY_SET_VALUE, &mut hkey).is_ok() {
            let _ = RegDeleteValueW(hkey, APP_NAME);
            let _ = RegCloseKey(hkey);
        }
    }
}

#[allow(dead_code)]
pub fn is_startup_enabled() -> bool {
    use windows::Win32::System::Registry::{RegQueryValueExW, KEY_READ};

    unsafe {
        let mut hkey: HKEY = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, STARTUP_KEY, 0, KEY_READ, &mut hkey).is_ok() {
            let result = RegQueryValueExW(hkey, APP_NAME, None, None, None, None).is_ok();
            let _ = RegCloseKey(hkey);
            return result;
        }
    }
    false
}
