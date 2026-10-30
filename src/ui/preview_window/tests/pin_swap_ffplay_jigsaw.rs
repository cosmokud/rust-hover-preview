use super::*;
use std::time::{Duration, Instant};

// The loop for the ffplay jigsaw: one pick from a portrait film to a
// landscape one, driven through the real swap while the real ffplay plays
// both, screenshotting what is on screen at every step.
//
// What it is for is seeing the scramble the fast case leaves — a single
// window whose picture is in tiles, only where the two shapes disagree, gone
// by the time the new film is playing — rather than reasoning about which
// window it could be. The log beside the pictures is the boxes every window
// standing in the switch has, read off the desktop: the pin's, the outgoing
// player's while its kill goes through, and the incoming one's from the
// moment the monitor names it.
//
// Env-gated like the native repro beside it: `RHP_REPRO_JIGSAW` names a
// folder holding `portrait.mp4` (a portrait film) and `landscape.mp4` (a
// landscape one). Without it the test reports what to set and returns. Make
// the pair with straight FFmpeg:
//
//   ffmpeg -f lavfi -i testsrc2=size=240x320:rate=30:duration=4 -pix_fmt yuv420p portrait.mp4
//   ffmpeg -f lavfi -i testsrc2=size=640x360:rate=30:duration=4 -pix_fmt yuv420p landscape.mp4
fn jigsaw_files() -> Option<(PathBuf, PathBuf)> {
    let dir = std::env::var("RHP_REPRO_JIGSAW").ok().map(PathBuf::from)?;
    let portrait = dir.join("portrait.mp4");
    let landscape = dir.join("landscape.mp4");
    (portrait.is_file() && landscape.is_file()).then_some((portrait, landscape))
}

/// The film boxes the switch is driven with: the portrait film standing where
/// a layout would put it, the landscape one at the same anchor — the way two
/// picks out of one listing land — so the two shapes overlap the way they do
/// on the desktop rather than side by side.
const PORTRAIT_BOX: (i32, i32, i32, i32) = (100, 80, 460, 560);
const LANDSCAPE_BOX: (i32, i32, i32, i32) = (100, 80, 1060, 620);

/// What the screenshots cover: both boxes with a margin for whatever stands
/// outside them.
const SHOT: (i32, i32, i32, i32) = (60, 40, 1060, 640);

/// A visible top-level window of `pid`, hwnd and box, read off the desktop.
struct SeenWindow {
    hwnd: isize,
    rect: (i32, i32, i32, i32),
}

struct Collect {
    pid: u32,
    found: Vec<SeenWindow>,
}

/// The `EnumWindows` callback the log reads the desktop with: every visible
/// top-level window the pid owns, largest first.
unsafe extern "system" fn collect_windows(
    hwnd: HWND,
    lparam: LPARAM,
) -> windows::Win32::Foundation::BOOL {
    use windows::Win32::Foundation::BOOL;
    let collect = &mut *(lparam.0 as *mut Collect);
    let mut window_pid: u32 = 0;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));
    if window_pid != collect.pid {
        return BOOL(1);
    }
    if !IsWindowVisible(hwnd).as_bool() {
        return BOOL(1);
    }
    if let Ok(owner) = GetWindow(hwnd, GW_OWNER) {
        if !owner.is_invalid() {
            return BOOL(1);
        }
    }
    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return BOOL(1);
    }
    if rect.right <= rect.left || rect.bottom <= rect.top {
        return BOOL(1);
    }
    collect.found.push(SeenWindow {
        hwnd: hwnd.0 as isize,
        rect: (rect.left, rect.top, rect.right, rect.bottom),
    });
    BOOL(1)
}

fn visible_windows_for_pid(pid: u32) -> Vec<SeenWindow> {
    if pid == 0 {
        return Vec::new();
    }
    let mut collect = Collect {
        pid,
        found: Vec::new(),
    };
    unsafe {
        let _ = EnumWindows(
            Some(collect_windows),
            LPARAM(&mut collect as *mut Collect as isize),
        );
    }
    collect.found.sort_by_key(|seen| {
        -((seen.rect.2 - seen.rect.0) as i64 * (seen.rect.3 - seen.rect.1) as i64)
    });
    collect.found
}

fn log_switch(tag: &str, pid_a: u32, hold_out: bool) {
    let pid_b = VIDEO_PID.load(Ordering::SeqCst);
    let hwnd = VIDEO_HWND.load(Ordering::SeqCst);
    let mut line =
        format!("jigsaw {tag}: hold_out={hold_out} VIDEO_PID={pid_b} VIDEO_HWND={hwnd:#x}");
    for (name, pid) in [("A", pid_a), ("B", pid_b)] {
        for seen in visible_windows_for_pid(pid) {
            let (l, t, r, b) = seen.rect;
            line.push_str(&format!(" {name}[hwnd={:#x} {l},{t},{r},{b}]", seen.hwnd));
        }
        if pid != 0 && visible_windows_for_pid(pid).is_empty() {
            line.push_str(&format!(" {name}[no visible window]"));
        }
    }
    eprintln!("{line}");
}

/// A screenshot of the switch, written beside the log so the boxes above can
/// be seen rather than imagined. BMP, bottom-up, the one encoding that needs
/// no crate: the desktop straight off the screen DC.
fn shoot(tag: &str) {
    let (x, y, width, height) = SHOT;
    let path = std::env::temp_dir()
        .join("jigsaw")
        .join(format!("jigsaw_{tag}.bmp"));
    let _ = std::fs::create_dir_all(path.parent().expect("a shot has a folder"));
    unsafe {
        use windows::Win32::Graphics::Gdi::*;
        let screen = GetDC(HWND(std::ptr::null_mut()));
        let mem = CreateCompatibleDC(screen);
        let bitmap = windows::Win32::Graphics::Gdi::CreateCompatibleBitmap(screen, width, height);
        let old = windows::Win32::Graphics::Gdi::SelectObject(mem, bitmap);
        let blitted =
            windows::Win32::Graphics::Gdi::BitBlt(mem, 0, 0, width, height, screen, x, y, SRCCOPY)
                .is_ok();
        windows::Win32::Graphics::Gdi::SelectObject(mem, old);
        let mut bits = vec![0u8; (width as usize) * (height as usize) * 4];
        let mut read = false;
        if blitted {
            let mut info = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: width,
                    biHeight: -height,
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: BI_RGB.0,
                    biSizeImage: 0,
                    biXPelsPerMeter: 0,
                    biYPelsPerMeter: 0,
                    biClrUsed: 0,
                    biClrImportant: 0,
                },
                bmiColors: [Default::default()],
            };
            read = windows::Win32::Graphics::Gdi::GetDIBits(
                mem,
                bitmap,
                0,
                height as u32,
                Some(bits.as_mut_ptr() as *mut core::ffi::c_void),
                &mut info,
                DIB_RGB_COLORS,
            ) > 0;
        }
        let _ = DeleteObject(bitmap);
        let _ = windows::Win32::Graphics::Gdi::DeleteDC(mem);
        let _ = windows::Win32::Graphics::Gdi::ReleaseDC(HWND(std::ptr::null_mut()), screen);

        if !read {
            eprintln!("jigsaw shot {tag}: the desktop would not be read");
            return;
        }
        // Bottom-up rows, BGRX order: the bytes GDI hands back are what a
        // viewer wants with the rows flipped.
        let stride = (width as usize) * 4;
        let mut file = vec![0u8; 54 + bits.len()];
        file[0] = b'B';
        file[1] = b'M';
        let size = file.len() as u32;
        file[2..6].copy_from_slice(&size.to_le_bytes());
        file[10] = 54;
        file[14..18].copy_from_slice(&40u32.to_le_bytes());
        file[18..22].copy_from_slice(&(width as u32).to_le_bytes());
        file[22..26].copy_from_slice(&(height as u32).to_le_bytes());
        file[26..28].copy_from_slice(&1u16.to_le_bytes());
        file[28..30].copy_from_slice(&32u16.to_le_bytes());
        for row in 0..(height as usize) {
            let from = row * stride;
            let to = 54 + (height as usize - 1 - row) * stride;
            file[to..to + stride].copy_from_slice(&bits[from..from + stride]);
        }
        if std::fs::write(&path, &file).is_ok() {
            eprintln!("jigsaw shot {tag}: {}", path.display());
        } else {
            eprintln!("jigsaw shot {tag}: unwritable");
        }
    }
}

/// The monitor's own role, played on cue rather than on its thread: find the
/// player's window by its id and style it the way the pass that found it
/// would, so `VIDEO_HWND` names what the app would have named and the window
/// stands in the order the app would have put it in.
fn monitor_pass(pid: u32) {
    if pid == 0 {
        return;
    }
    unsafe {
        let _ = try_apply_noactivate_style(pid);
    }
}

fn wait_for_player_window(pid: u32, what: &str, deadline: Duration) -> bool {
    let started = Instant::now();
    while started.elapsed() < deadline {
        if !visible_windows_for_pid(pid).is_empty() {
            return true;
        }
        std::thread::sleep(Duration::from_millis(50));
    }
    eprintln!(
        "jigsaw: NO window for {what} (pid={pid}) in {} ms",
        started.elapsed().as_millis()
    );
    false
}

#[test]
fn a_portrait_to_landscape_pick_on_ffplay() {
    let Some((portrait, landscape)) = jigsaw_files() else {
        println!(
            "set RHP_REPRO_JIGSAW to a folder holding portrait.mp4 (portrait) and \
             landscape.mp4 (landscape) to run this"
        );
        return;
    };
    let _room = the_room_to_itself();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    // Real players, really spawned: the noting every other test runs on begins
    // nothing, and a loop watching windows come and go cannot run on noted ones.
    spawn_real_player(true);

    let mut failure: Option<String> = None;

    // The probe the launch reads its box and its crop out of, the way a
    // hover's pass primes it before any pick exists.
    match probe_video_geometry(&portrait) {
        ProbedGeometry::Measured(g) => {
            eprintln!("jigsaw: portrait probe measured {}x{}", g.width, g.height)
        }
        ProbedGeometry::Unmeasurable => eprintln!("jigsaw: portrait probe UNMEASURABLE"),
    }
    match probe_video_geometry(&landscape) {
        ProbedGeometry::Measured(g) => {
            eprintln!("jigsaw: landscape probe measured {}x{}", g.width, g.height)
        }
        ProbedGeometry::Unmeasurable => eprintln!("jigsaw: landscape probe UNMEASURABLE"),
    }

    // The pin standing on the portrait film, and the portrait film playing
    // in it: the player begun the way the take-up begins one, at the box the
    // layout put the film in, silent the way a test is.
    let (ax, ay, ar, ab) = PORTRAIT_BOX;
    stand_pin(Some(overlay_pin((ax, ay, ar, ab), PinChrome::always())));
    let mut standing = create_loading_media(360, 480);
    standing.media_type = MediaType::Video;
    let process_a = start_video_playback(
        &portrait,
        ax,
        ay,
        ar - ax,
        ab - ay,
        0.0,
        false,
        0,
        None,
        None,
    );
    let pid_a = process_a.as_ref().map(PlayerProcess::id).unwrap_or(0);
    eprintln!("jigsaw: portrait ffplay pid={pid_a}");
    if process_a.is_none() {
        failure = Some("the portrait film's player would not begin".to_string());
    } else {
        standing.video_process = process_a;
    }
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(standing);
    }
    let mut pid_b: u32 = 0;

    if failure.is_none() {
        if !wait_for_player_window(pid_a, "portrait", Duration::from_secs(10)) {
            failure = Some("the portrait player never put a window up".to_string());
        } else {
            monitor_pass(pid_a);
            // The first frame takes a decode: what the shot shows before it
            // is the window ffplay opens, and after it the film.
            std::thread::sleep(Duration::from_millis(800));
            log_switch("a_playing", pid_a, false);
            shoot("a_playing");
        }
    }

    // The pick: the landscape file through the real swap — the cover over
    // the outgoing frame, the take-down, the kill, the launch, and the
    // hold for the incoming window.
    if failure.is_none() {
        let (bx, by, br, bb) = LANDSCAPE_BOX;
        let mut media = create_loading_media((br - bx) as u32, (bb - by) as u32);
        media.media_type = MediaType::Video;
        let swap = swap_pinned_media(PinAnswer {
            path: landscape.clone(),
            update: PinUpdate {
                content: (bx, by, br, bb),
                dpi: 96,
                volume: 0,
            },
            media: Some(media),
            walk: None,
            arc: PinArc::new(),
        });
        let mut hold = None;
        abandon_pin_swap(&mut hold);
        match swap {
            PinSwap::Holding(held) => {
                hold = Some(held);
                eprintln!("jigsaw: the pick is held for the landscape window");
            }
            PinSwap::Ready(_) => {
                failure = Some("the pick installed at once with a cover standing".to_string());
            }
            PinSwap::Refused { .. } => {
                failure = Some("the pick refused the landscape film".to_string());
            }
        }
        pid_b = VIDEO_PID.load(Ordering::SeqCst);
        eprintln!("jigsaw: landscape ffplay pid={pid_b}");
        log_switch("after_pick", pid_a, true);
        shoot("after_pick");

        // The ticks the switch lives through: the monitor's pass and the
        // hold's, with the desktop read after every one — until the file is
        // installed or the switch runs out of reasons to go on.
        let mut load = None;
        let settled = Instant::now();
        let mut shot = 0;
        let mut installed = false;
        let mut after_install = 0;
        while settled.elapsed() < Duration::from_secs(8) {
            monitor_pass(VIDEO_PID.load(Ordering::SeqCst));
            // The tick's own place for a film whose band has moved: the handover tick puts
            // the window at the box the pin stood at, and this is what carries it to the
            // file's own on the ticks after (see `ensure_pinned_sibling_box`).
            if installed {
                ensure_pinned_sibling_box(LANDSCAPE_BOX);
            }
            match settle_pin_swap(&mut hold, &mut load) {
                Some(PinSwap::Ready(_)) => {
                    installed = true;
                }
                Some(PinSwap::Refused { .. }) => {
                    failure = Some("the held landscape film was refused".to_string());
                    break;
                }
                Some(PinSwap::Holding(held)) => {
                    hold = Some(held);
                }
                None => {}
            }
            shot += 1;
            log_switch(&format!("settle_{shot}"), pid_a, hold.is_some());
            shoot(&format!("settle_{shot:02}"));
            if installed {
                after_install += 1;
                // The install the tick performs before the park is settled: the pin takes up
                // the incoming file's box, so the handover below puts the window where the
                // file stands rather than where the film it replaces stood.
                if after_install == 1 {
                    if let Some(mut pinned) = pin_state() {
                        if let Some(pin) = pinned.pin_mut() {
                            pin.content = LANDSCAPE_BOX;
                        }
                    }
                }
                if after_install >= 5 {
                    break;
                }
            }
            std::thread::sleep(Duration::from_millis(50));
        }
        log_switch("end", pid_a, hold.is_some());
        if failure.is_none() && !installed {
            failure = Some("the hold never settled in 8 s".to_string());
        }
        if let Some(held) = hold.take() {
            abandon_pin_swap(&mut Some(held));
        }
    }

    // Whatever is still standing is ended the way the sweep ends it, by id
    // and under the name it was begun with — and the globals go back to
    // what the next test walked into.
    for pid in [pid_a, pid_b] {
        if pid != 0 && is_player_pid_alive(pid) {
            terminate_player_pid(pid);
        }
    }
    clear_pending_pinned_relaunch();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    spawn_real_player(false);

    assert!(failure.is_none(), "jigsaw loop failed: {failure:?}");
}
