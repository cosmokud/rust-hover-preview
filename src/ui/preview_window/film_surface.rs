//! Owned `FilmSurface` seam and `FilmHostSession` state machine.
//!
//! Consolidates film host HWND creation/destruction, DWM thumbnail registration/update/teardown,
//! resize, display/DPI move, and chrome ownership sync behind a single testable seam.
//! RAII guards (`HostGuard`, `ThumbnailGuard`) guarantee 1:1 create/destroy and
//! register/unregister parity across normal teardown, swap, and error/early-return paths.

use std::path::PathBuf;

use windows::Win32::Foundation::{HWND, RECT};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_RECTSOURCE,
    DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::UI::WindowsAndMessaging::{SetWindowPos, SWP_NOACTIVATE, SWP_NOZORDER};

use super::pin_trace::note_pin_trace;
use super::{pin_film_host, ScreenRegion};

/// Platform seam for film host HWNDs and their DWM thumbnail registrations.
/// Handles are `isize` (`0` = invalid/none) so headless tests can run without Win32 windows.
pub(super) trait FilmSurface {
    fn create_host(&self, always_on_top: bool) -> Option<isize>;
    fn destroy_host(&self, host: isize) -> bool;
    fn register_thumbnail(&self, host: isize, source: isize) -> Option<isize>;
    fn unregister_thumbnail(&self, registration: isize) -> bool;
    fn update_thumbnail(
        &self,
        registration: isize,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    ) -> bool;
    fn resize_host(&self, host: isize, region: ScreenRegion) -> bool;
    fn move_display(&self, host: isize, region: ScreenRegion, dpi: u32) -> bool;
    fn sync_chrome(&self, host: isize, owned: bool, visible: bool) -> bool;
}

/// RAII guard for a newly created host HWND. Destroys on drop unless committed.
pub(super) struct HostGuard<'a> {
    surface: &'a dyn FilmSurface,
    hwnd: isize,
    armed: bool,
}

impl<'a> HostGuard<'a> {
    pub(super) fn create(surface: &'a dyn FilmSurface, always_on_top: bool) -> Option<Self> {
        let hwnd = surface.create_host(always_on_top)?;
        if hwnd == 0 {
            return None;
        }
        Some(Self {
            surface,
            hwnd,
            armed: true,
        })
    }

    pub(super) fn hwnd(&self) -> isize {
        self.hwnd
    }

    pub(super) fn commit(mut self) -> isize {
        self.armed = false;
        self.hwnd
    }
}

impl Drop for HostGuard<'_> {
    fn drop(&mut self) {
        if self.armed && self.hwnd != 0 {
            let _ = self.surface.destroy_host(self.hwnd);
        }
    }
}

/// RAII guard for a newly registered DWM thumbnail. Unregisters on drop unless committed.
pub(super) struct ThumbnailGuard<'a> {
    surface: &'a dyn FilmSurface,
    registration: isize,
    armed: bool,
}

impl<'a> ThumbnailGuard<'a> {
    pub(super) fn register(
        surface: &'a dyn FilmSurface,
        host: isize,
        source: isize,
    ) -> Option<Self> {
        let registration = surface.register_thumbnail(host, source)?;
        if registration == 0 {
            return None;
        }
        Some(Self {
            surface,
            registration,
            armed: true,
        })
    }

    pub(super) fn registration(&self) -> isize {
        self.registration
    }

    pub(super) fn commit(mut self) -> isize {
        self.armed = false;
        self.registration
    }
}

impl Drop for ThumbnailGuard<'_> {
    fn drop(&mut self) {
        if self.armed && self.registration != 0 {
            let _ = self.surface.unregister_thumbnail(self.registration);
        }
    }
}

/// Active DWM thumbnail binding owned by a [`FilmHostSession`].
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) struct ActiveThumbnail {
    pub(super) registration: isize,
    pub(super) host: isize,
    pub(super) source: isize,
    pub(super) source_process: u32,
    pub(super) dest: (i32, i32, i32, i32),
    pub(super) src: Option<(i32, i32, i32, i32)>,
}

/// Cohesive owned state for a single film host lifecycle (attach, resize, move-display,
/// sync-chrome, source swap, and unified teardown).
#[derive(Default)]
pub(super) struct FilmHostSession {
    host_hwnd: Option<isize>,
    player_pid: Option<u32>,
    player_path: Option<PathBuf>,
    presented_pid: u32,
    thumbnail: Option<ActiveThumbnail>,
    chrome_owned: bool,
    visible: bool,
    region: Option<ScreenRegion>,
    dpi: u32,
}

#[allow(dead_code)]
impl FilmHostSession {
    pub(super) fn new() -> Self {
        Self::default()
    }

    pub(super) fn host(&self) -> Option<isize> {
        self.host_hwnd
    }

    pub(super) fn thumbnail(&self) -> Option<ActiveThumbnail> {
        self.thumbnail
    }

    pub(super) fn player_pid(&self) -> Option<u32> {
        self.player_pid
    }

    pub(super) fn set_player(&mut self, pid: Option<u32>, path: Option<PathBuf>) {
        self.player_pid = pid;
        self.player_path = path;
    }

    pub(super) fn presented_pid(&self) -> u32 {
        self.presented_pid
    }

    pub(super) fn set_presented_pid(&mut self, pid: u32) {
        self.presented_pid = pid;
    }

    pub(super) fn region(&self) -> Option<ScreenRegion> {
        self.region
    }

    pub(super) fn dpi(&self) -> u32 {
        self.dpi
    }

    pub(super) fn chrome_owned(&self) -> bool {
        self.chrome_owned
    }

    pub(super) fn visible(&self) -> bool {
        self.visible
    }

    /// Attach a host HWND (either adopting an existing hover host or creating a fresh one).
    /// Tears down any previously owned host first so no HWND is ever overwritten/leaked.
    pub(super) fn attach(
        &mut self,
        surface: &dyn FilmSurface,
        adopted_host: Option<isize>,
        always_on_top: bool,
    ) -> Option<isize> {
        let adopted = adopted_host.filter(|&h| h != 0);
        if let Some(existing) = self.host_hwnd {
            if Some(existing) == adopted {
                return Some(existing);
            }
            self.detach(surface);
        }

        if let Some(host) = adopted {
            if !surface.sync_chrome(host, true, false) {
                return None;
            }
            self.host_hwnd = Some(host);
            self.chrome_owned = true;
            self.visible = false;
            return Some(host);
        }

        let guard = HostGuard::create(surface, always_on_top)?;
        let hwnd = guard.hwnd();
        if !surface.sync_chrome(hwnd, true, false) {
            // HostGuard drops and destroys the freshly created HWND.
            return None;
        }
        let hwnd = guard.commit();
        self.host_hwnd = Some(hwnd);
        self.chrome_owned = true;
        self.visible = false;
        Some(hwnd)
    }

    /// Register and configure a DWM thumbnail for `source` in the attached host.
    /// Rolls back the registration if property update fails.
    pub(super) fn attach_source(
        &mut self,
        surface: &dyn FilmSurface,
        source: isize,
        pid: u32,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    ) -> bool {
        let Some(host) = self.host_hwnd else {
            return false;
        };
        if source == 0 {
            return false;
        }

        // Fast path: same host & source already registered -> update in place.
        if let Some(mut active) = self.thumbnail {
            if active.host == host && active.source == source {
                if !surface.update_thumbnail(active.registration, dest, src) {
                    // Failed update invalidates the standing registration; unregister to keep parity.
                    let _ = surface.unregister_thumbnail(active.registration);
                    self.thumbnail = None;
                    return false;
                }
                active.source_process = pid;
                active.dest = dest;
                active.src = src;
                self.thumbnail = Some(active);
                return true;
            }
        }

        self.swap_source(surface, source, pid, dest, src)
    }

    /// Hand over the thumbnail to a new source window without blanking the old one on failure.
    /// Guarantees 1:1 register/unregister parity on both success and error paths.
    pub(super) fn swap_source(
        &mut self,
        surface: &dyn FilmSurface,
        new_source: isize,
        new_pid: u32,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    ) -> bool {
        let Some(host) = self.host_hwnd else {
            return false;
        };
        if new_source == 0 {
            return false;
        }

        let Some(guard) = ThumbnailGuard::register(surface, host, new_source) else {
            return false;
        };
        if !surface.update_thumbnail(guard.registration(), dest, src) {
            // ThumbnailGuard drops -> unregisters incoming registration; old thumbnail stays intact.
            return false;
        }
        let registration = guard.commit();
        let previous = self.thumbnail.replace(ActiveThumbnail {
            registration,
            host,
            source: new_source,
            source_process: new_pid,
            dest,
            src,
        });
        if let Some(old) = previous {
            let _ = surface.unregister_thumbnail(old.registration);
        }
        true
    }

    /// Resize the host and re-point the active thumbnail to the new destination/crop.
    pub(super) fn resize(
        &mut self,
        surface: &dyn FilmSurface,
        region: ScreenRegion,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    ) -> bool {
        let Some(host) = self.host_hwnd else {
            return false;
        };
        if !surface.resize_host(host, region) {
            return false;
        }
        self.region = Some(region);
        if let Some(mut active) = self.thumbnail {
            if !surface.update_thumbnail(active.registration, dest, src) {
                let _ = surface.unregister_thumbnail(active.registration);
                self.thumbnail = None;
                return false;
            }
            active.dest = dest;
            active.src = src;
            self.thumbnail = Some(active);
        }
        true
    }

    /// Move the host onto a display with a new DPI and update the active thumbnail.
    pub(super) fn move_display(
        &mut self,
        surface: &dyn FilmSurface,
        region: ScreenRegion,
        dpi: u32,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    ) -> bool {
        let Some(host) = self.host_hwnd else {
            return false;
        };
        if !surface.move_display(host, region, dpi) {
            return false;
        }
        self.region = Some(region);
        self.dpi = dpi;
        if let Some(mut active) = self.thumbnail {
            if !surface.update_thumbnail(active.registration, dest, src) {
                let _ = surface.unregister_thumbnail(active.registration);
                self.thumbnail = None;
                return false;
            }
            active.dest = dest;
            active.src = src;
            self.thumbnail = Some(active);
        }
        true
    }

    /// Synchronize pin chrome ownership and host visibility.
    pub(super) fn sync_chrome(
        &mut self,
        surface: &dyn FilmSurface,
        owned: bool,
        visible: bool,
    ) -> bool {
        let Some(host) = self.host_hwnd else {
            return false;
        };
        if !surface.sync_chrome(host, owned, visible) {
            return false;
        }
        self.chrome_owned = owned;
        self.visible = visible;
        true
    }

    /// Unified teardown: unregister thumbnail -> unown chrome -> destroy host -> clear state.
    /// Order is load-bearing: DWM thumbnail must unregister before its destination HWND dies,
    /// and pin chrome must be unowned before destroying the host or Windows destroys the chrome.
    pub(super) fn detach(&mut self, surface: &dyn FilmSurface) {
        if let Some(active) = self.thumbnail.take() {
            let _ = surface.unregister_thumbnail(active.registration);
        }
        if let Some(host) = self.host_hwnd.take() {
            if self.chrome_owned {
                let _ = surface.sync_chrome(host, false, false);
            }
            let _ = surface.destroy_host(host);
        }
        self.player_pid = None;
        self.player_path = None;
        self.presented_pid = 0;
        self.chrome_owned = false;
        self.visible = false;
        self.region = None;
        self.dpi = 0;
    }
}

/// Real Win32/DWM adapter for [`FilmSurface`].
#[allow(dead_code)]
#[derive(Clone, Copy)]
pub(super) struct Win32FilmSurface;

impl FilmSurface for Win32FilmSurface {
    fn create_host(&self, always_on_top: bool) -> Option<isize> {
        unsafe { pin_film_host::create_raw_host(always_on_top) }.map(|hwnd| hwnd.0 as isize)
    }

    fn destroy_host(&self, host: isize) -> bool {
        if host == 0 {
            return false;
        }
        unsafe { pin_film_host::destroy_raw_host(HWND(host as *mut _)) }
    }

    fn register_thumbnail(&self, host: isize, source: isize) -> Option<isize> {
        if host == 0 || source == 0 {
            return None;
        }
        match unsafe { DwmRegisterThumbnail(HWND(host as *mut _), HWND(source as *mut _)) } {
            Ok(registration) => Some(registration),
            Err(err) => {
                note_pin_trace!(
                    "film surface DwmRegisterThumbnail failed host {host:#x} source {source:#x}: {err:?}"
                );
                None
            }
        }
    }

    fn unregister_thumbnail(&self, registration: isize) -> bool {
        if registration == 0 {
            return false;
        }
        match unsafe { DwmUnregisterThumbnail(registration) } {
            Ok(()) => true,
            Err(err) => {
                note_pin_trace!(
                    "film surface DwmUnregisterThumbnail failed reg {registration:#x}: {err:?}"
                );
                false
            }
        }
    }

    fn update_thumbnail(
        &self,
        registration: isize,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    ) -> bool {
        if registration == 0 {
            return false;
        }
        let (x, y, width, height) = dest;
        let mut props = DWM_THUMBNAIL_PROPERTIES {
            dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE,
            rcDestination: RECT {
                left: x,
                top: y,
                right: x + width,
                bottom: y + height,
            },
            rcSource: RECT::default(),
            opacity: 255,
            fVisible: true.into(),
            fSourceClientAreaOnly: true.into(),
        };
        if let Some((left, top, right, bottom)) = src {
            props.dwFlags |= DWM_TNP_RECTSOURCE;
            props.rcSource = RECT {
                left,
                top,
                right,
                bottom,
            };
        }
        match unsafe { DwmUpdateThumbnailProperties(registration, &props) } {
            Ok(()) => true,
            Err(err) => {
                note_pin_trace!(
                    "film surface DwmUpdateThumbnailProperties failed reg {registration:#x}: {err:?}"
                );
                false
            }
        }
    }

    fn resize_host(&self, host: isize, region: ScreenRegion) -> bool {
        if host == 0 {
            return false;
        }
        let (left, top, right, bottom) = region;
        match unsafe {
            SetWindowPos(
                HWND(host as *mut _),
                HWND::default(),
                left,
                top,
                (right - left).max(1),
                (bottom - top).max(1),
                SWP_NOZORDER | SWP_NOACTIVATE,
            )
        } {
            Ok(()) => true,
            Err(err) => {
                note_pin_trace!("film surface resize_host failed host {host:#x}: {err:?}");
                false
            }
        }
    }

    fn move_display(&self, host: isize, region: ScreenRegion, _dpi: u32) -> bool {
        self.resize_host(host, region)
    }

    fn sync_chrome(&self, _host: isize, owned: bool, visible: bool) -> bool {
        unsafe {
            pin_film_host::sync_chrome_ownership(owned);
            pin_film_host::set_the_pin_film_host_visible(visible);
        }
        true
    }
}

#[cfg(test)]
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum FilmSurfaceCall {
    CreateHost {
        always_on_top: bool,
        hwnd: isize,
    },
    DestroyHost(isize),
    RegisterThumbnail {
        host: isize,
        source: isize,
        registration: isize,
    },
    UnregisterThumbnail(isize),
    UpdateThumbnail {
        registration: isize,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    },
    ResizeHost {
        host: isize,
        region: ScreenRegion,
    },
    MoveDisplay {
        host: isize,
        region: ScreenRegion,
        dpi: u32,
    },
    SyncChrome {
        host: isize,
        owned: bool,
        visible: bool,
    },
}

#[cfg(test)]
pub(super) struct RecordedFilmSurface {
    state: std::sync::Mutex<RecordedFilmState>,
}

#[cfg(test)]
struct RecordedFilmState {
    next_hwnd: isize,
    next_reg: isize,
    calls: Vec<FilmSurfaceCall>,
    live_hosts: Vec<isize>,
    live_thumbnails: Vec<isize>,
    created_hosts: u32,
    destroyed_hosts: u32,
    registered_thumbnails: u32,
    unregistered_thumbnails: u32,
    fail_create_host: bool,
    fail_register_thumbnail: bool,
    fail_update_thumbnail: bool,
    fail_sync_chrome: bool,
    fail_resize_host: bool,
}

#[cfg(test)]
impl RecordedFilmSurface {
    pub(super) fn new() -> Self {
        Self {
            state: std::sync::Mutex::new(RecordedFilmState {
                next_hwnd: 0x5000,
                next_reg: 0x9000,
                calls: Vec::new(),
                live_hosts: Vec::new(),
                live_thumbnails: Vec::new(),
                created_hosts: 0,
                destroyed_hosts: 0,
                registered_thumbnails: 0,
                unregistered_thumbnails: 0,
                fail_create_host: false,
                fail_register_thumbnail: false,
                fail_update_thumbnail: false,
                fail_sync_chrome: false,
                fail_resize_host: false,
            }),
        }
    }

    pub(super) fn seed_adopted_host(&self, hwnd: isize) {
        let mut s = self.state.lock().unwrap();
        s.created_hosts += 1;
        s.live_hosts.push(hwnd);
    }

    pub(super) fn fail_create_host(&self, fail: bool) {
        self.state.lock().unwrap().fail_create_host = fail;
    }

    pub(super) fn fail_register_thumbnail(&self, fail: bool) {
        self.state.lock().unwrap().fail_register_thumbnail = fail;
    }

    pub(super) fn fail_update_thumbnail(&self, fail: bool) {
        self.state.lock().unwrap().fail_update_thumbnail = fail;
    }

    pub(super) fn fail_sync_chrome(&self, fail: bool) {
        self.state.lock().unwrap().fail_sync_chrome = fail;
    }

    pub(super) fn calls(&self) -> Vec<FilmSurfaceCall> {
        self.state.lock().unwrap().calls.clone()
    }

    pub(super) fn host_counts(&self) -> (u32, u32) {
        let s = self.state.lock().unwrap();
        (s.created_hosts, s.destroyed_hosts)
    }

    pub(super) fn thumbnail_counts(&self) -> (u32, u32) {
        let s = self.state.lock().unwrap();
        (s.registered_thumbnails, s.unregistered_thumbnails)
    }

    pub(super) fn live_hosts(&self) -> Vec<isize> {
        self.state.lock().unwrap().live_hosts.clone()
    }

    pub(super) fn live_thumbnails(&self) -> Vec<isize> {
        self.state.lock().unwrap().live_thumbnails.clone()
    }

    pub(super) fn has_no_leaks(&self) -> bool {
        let s = self.state.lock().unwrap();
        s.live_hosts.is_empty()
            && s.live_thumbnails.is_empty()
            && s.created_hosts == s.destroyed_hosts
            && s.registered_thumbnails == s.unregistered_thumbnails
    }
}

#[cfg(test)]
impl FilmSurface for RecordedFilmSurface {
    fn create_host(&self, always_on_top: bool) -> Option<isize> {
        let mut s = self.state.lock().unwrap();
        if s.fail_create_host {
            return None;
        }
        let hwnd = s.next_hwnd;
        s.next_hwnd += 0x10;
        s.created_hosts += 1;
        s.live_hosts.push(hwnd);
        s.calls.push(FilmSurfaceCall::CreateHost {
            always_on_top,
            hwnd,
        });
        Some(hwnd)
    }

    fn destroy_host(&self, host: isize) -> bool {
        let mut s = self.state.lock().unwrap();
        s.calls.push(FilmSurfaceCall::DestroyHost(host));
        if let Some(pos) = s.live_hosts.iter().position(|&h| h == host) {
            s.live_hosts.remove(pos);
            s.destroyed_hosts += 1;
            true
        } else {
            false
        }
    }

    fn register_thumbnail(&self, host: isize, source: isize) -> Option<isize> {
        let mut s = self.state.lock().unwrap();
        if s.fail_register_thumbnail {
            return None;
        }
        let registration = s.next_reg;
        s.next_reg += 0x10;
        s.registered_thumbnails += 1;
        s.live_thumbnails.push(registration);
        s.calls.push(FilmSurfaceCall::RegisterThumbnail {
            host,
            source,
            registration,
        });
        Some(registration)
    }

    fn unregister_thumbnail(&self, registration: isize) -> bool {
        let mut s = self.state.lock().unwrap();
        s.calls
            .push(FilmSurfaceCall::UnregisterThumbnail(registration));
        if let Some(pos) = s.live_thumbnails.iter().position(|&r| r == registration) {
            s.live_thumbnails.remove(pos);
            s.unregistered_thumbnails += 1;
            true
        } else {
            false
        }
    }

    fn update_thumbnail(
        &self,
        registration: isize,
        dest: (i32, i32, i32, i32),
        src: Option<(i32, i32, i32, i32)>,
    ) -> bool {
        let mut s = self.state.lock().unwrap();
        s.calls.push(FilmSurfaceCall::UpdateThumbnail {
            registration,
            dest,
            src,
        });
        !s.fail_update_thumbnail
    }

    fn resize_host(&self, host: isize, region: ScreenRegion) -> bool {
        let mut s = self.state.lock().unwrap();
        s.calls.push(FilmSurfaceCall::ResizeHost { host, region });
        !s.fail_resize_host
    }

    fn move_display(&self, host: isize, region: ScreenRegion, dpi: u32) -> bool {
        let mut s = self.state.lock().unwrap();
        s.calls
            .push(FilmSurfaceCall::MoveDisplay { host, region, dpi });
        !s.fail_resize_host
    }

    fn sync_chrome(&self, host: isize, owned: bool, visible: bool) -> bool {
        let mut s = self.state.lock().unwrap();
        s.calls.push(FilmSurfaceCall::SyncChrome {
            host,
            owned,
            visible,
        });
        !s.fail_sync_chrome
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normal_lifecycle_maintains_hwnd_and_thumbnail_parity() {
        let surface = RecordedFilmSurface::new();
        let mut session = FilmHostSession::new();

        let host = session.attach(&surface, None, true).expect("host created");
        assert_eq!(session.host(), Some(host));
        assert!(session.chrome_owned());
        assert!(!session.visible());

        assert!(session.attach_source(&surface, 0x7001, 42, (0, 0, 800, 450), None));
        assert!(session.resize(
            &surface,
            (100, 100, 1060, 640),
            (0, 0, 960, 540),
            Some((0, 0, 1920, 1080)),
        ));
        assert_eq!(session.region(), Some((100, 100, 1060, 640)));

        assert!(session.move_display(
            &surface,
            (2000, 100, 3200, 775),
            144,
            (0, 0, 1200, 675),
            Some((0, 0, 1920, 1080)),
        ));
        assert_eq!(session.dpi(), 144);
        assert!(session.sync_chrome(&surface, true, true));
        assert!(session.visible());

        session.detach(&surface);
        // Idempotent second detach does not double-destroy.
        session.detach(&surface);

        assert!(surface.has_no_leaks());
        assert_eq!(surface.host_counts(), (1, 1));
        assert_eq!(surface.thumbnail_counts(), (1, 1));

        // Teardown order: unregister thumbnail -> unown chrome -> destroy host.
        let calls = surface.calls();
        let unreg_idx = calls
            .iter()
            .position(|c| matches!(c, FilmSurfaceCall::UnregisterThumbnail(_)))
            .unwrap();
        let unown_idx = calls
            .iter()
            .position(|c| {
                matches!(
                    c,
                    FilmSurfaceCall::SyncChrome {
                        owned: false,
                        visible: false,
                        ..
                    }
                )
            })
            .unwrap();
        let destroy_idx = calls
            .iter()
            .position(|c| matches!(c, FilmSurfaceCall::DestroyHost(_)))
            .unwrap();
        assert!(unreg_idx < unown_idx);
        assert!(unown_idx < destroy_idx);
    }

    #[test]
    fn swap_source_maintains_parity_on_success_and_preserves_old_on_error() {
        let surface = RecordedFilmSurface::new();
        let mut session = FilmHostSession::new();
        session.attach(&surface, None, false).unwrap();

        assert!(session.attach_source(&surface, 0x7001, 10, (0, 0, 640, 360), None));
        let first_reg = session.thumbnail().unwrap().registration;

        // Failed update on incoming swap rolls back incoming thumbnail and keeps first_reg alive.
        surface.fail_update_thumbnail(true);
        assert!(!session.swap_source(&surface, 0x7002, 20, (0, 0, 640, 360), None));
        assert_eq!(session.thumbnail().unwrap().registration, first_reg);
        assert_eq!(surface.live_thumbnails(), vec![first_reg]);
        assert_eq!(surface.thumbnail_counts(), (2, 1));

        // Successful swap unregisters first_reg and keeps only the new one.
        surface.fail_update_thumbnail(false);
        assert!(session.swap_source(&surface, 0x7002, 20, (0, 0, 640, 360), None));
        let second_reg = session.thumbnail().unwrap().registration;
        assert_ne!(first_reg, second_reg);
        assert_eq!(surface.live_thumbnails(), vec![second_reg]);
        assert_eq!(surface.thumbnail_counts(), (3, 2));

        session.detach(&surface);
        assert!(surface.has_no_leaks());
        assert_eq!(surface.host_counts(), (1, 1));
        assert_eq!(surface.thumbnail_counts(), (3, 3));
    }

    #[test]
    fn error_and_early_return_paths_never_leak_hwnd_or_thumbnail() {
        let surface = RecordedFilmSurface::new();
        let mut session = FilmHostSession::new();

        // Create-host failure leaves no host and no leaks.
        surface.fail_create_host(true);
        assert!(session.attach(&surface, None, true).is_none());
        assert!(session.host().is_none());
        assert!(surface.has_no_leaks());
        surface.fail_create_host(false);

        // Sync-chrome failure during attach must roll back the freshly created host HWND.
        surface.fail_sync_chrome(true);
        assert!(session.attach(&surface, None, true).is_none());
        assert!(session.host().is_none());
        assert!(surface.has_no_leaks());
        assert_eq!(surface.host_counts(), (1, 1));

        // Recover and attach cleanly.
        surface.fail_sync_chrome(false);
        let first_host = session.attach(&surface, None, true).unwrap();
        assert!(session.attach_source(&surface, 0x7001, 10, (0, 0, 640, 360), None));

        // Adopting a different host tears down the previous host + thumbnail first.
        surface.seed_adopted_host(0x6000);
        let adopted = session.attach(&surface, Some(0x6000), false).unwrap();
        assert_eq!(adopted, 0x6000);
        assert_eq!(surface.live_hosts(), vec![0x6000]);
        assert!(surface.live_thumbnails().is_empty());
        assert_ne!(first_host, adopted);

        // Register failure leaves no leaks.
        surface.fail_register_thumbnail(true);
        assert!(!session.attach_source(&surface, 0x7002, 20, (0, 0, 640, 360), None));

        session.detach(&surface);
        assert!(surface.has_no_leaks());
    }

    #[test]
    fn raii_guards_clean_up_on_early_return() {
        let surface = RecordedFilmSurface::new();
        {
            let host = HostGuard::create(&surface, true).unwrap();
            let _thumb = ThumbnailGuard::register(&surface, host.hwnd(), 0x8001).unwrap();
            // Early return without commit(): both guards drop in LIFO order (thumb then host).
        }
        assert!(surface.has_no_leaks());
        assert_eq!(surface.host_counts(), (1, 1));
        assert_eq!(surface.thumbnail_counts(), (1, 1));
    }
}
