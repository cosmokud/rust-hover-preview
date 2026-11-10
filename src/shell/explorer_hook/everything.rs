//! voidtools **Everything**'s result list: the second listing that is nobody's Shell window, read
//! out of Everything's own address space.
//!
//! **Nothing about this listing is a position the rest of the hook can ask for.** Everything's frame
//! is not a Shell window, so `Shell.Windows()` has nothing for it, no Shell view is registered for
//! it, and its rows are a plain `SysListView32` of another process whose items carry no `ItemIndex`
//! and no path. A search also spans folders, so there is no folder a name could be resolved against
//! either: the row's own cells are the only witness there is, which is why they are the whole of what
//! this module reads (see `row_file`).
//!
//! **And they are never read through UI Automation.** Reading that list through UIA is reported to
//! hang Everything and to leave it holding gigabytes, so every road that could reach
//! `ElementFromPoint`/`GetFocusedElement` asks the frame's class first and comes here instead — see
//! `uia_item_from_point` and `uia_item_from_focus`, which is where every walk starts.
//!
//! What replaces the walk is the list's own messages. Everything answers `LVM_*` itself (its list
//! view is subclassed), and every one of them that carries a struct reads that struct out of the
//! *caller's* address space — so a struct handed to it has to be in Everything's process: `OpenProcess`,
//! one 4 KiB scratch page in that process, `WriteProcessMemory` the struct, `SendMessageTimeout`, and
//! `ReadProcessMemory` the answer back. Every message sent here is a read — a count, a hit test, a
//! row's box, a cell's text — and nothing that changes what Everything holds is ever sent.
//!
//! What it costs is measured, and it is why nothing is memoized here: one full read — a hit test and
//! two cells — is about 0.1 ms against a 135 000-row list, a fraction of the walk it stands in for.
//! The memo the pointer road keeps is the one thing that would be wrong here rather than slow: a
//! query typed into Everything's search box puts another file at the same box under a parked pointer
//! (see `get_file_under_cursor`).
//!
//! Everything running elevated answers nothing: `OpenProcess` fails under UIPI and every message to
//! its window is blocked. That is an absence rather than an error — a page that could not be made is
//! left alone for a while rather than asked for on every tick, and nothing about it is logged.

use crate::shell::window_classes::utf16_contains_ascii_ignore_case;
use windows::core::PWSTR;
use windows::Win32::Foundation::{CloseHandle, HANDLE, WPARAM};
use windows::Win32::Graphics::Gdi::{ClientToScreen, ScreenToClient};
use windows::Win32::System::Diagnostics::Debug::{ReadProcessMemory, WriteProcessMemory};
use windows::Win32::System::Memory::{
    VirtualAllocEx, VirtualFreeEx, MEM_COMMIT, MEM_RELEASE, MEM_RESERVE, PAGE_READWRITE,
};
use windows::Win32::System::Threading::{
    OpenProcess, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_VM_OPERATION, PROCESS_VM_READ,
    PROCESS_VM_WRITE,
};
use windows::Win32::UI::Controls::{
    HDM_GETITEMCOUNT, LVHITTESTINFO, LVIF_TEXT, LVIR_BOUNDS, LVITEMW, LVM_GETHEADER,
    LVM_GETITEMCOUNT, LVM_GETITEMRECT, LVM_GETITEMTEXTW, LVM_GETNEXTITEM, LVM_GETTOPINDEX,
    LVM_HITTEST, LVNI_FOCUSED, LVNI_SELECTED,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumChildWindows, GetClientRect, SendMessageTimeoutW, SMTO_ABORTIFHUNG, SMTO_BLOCK,
};

use super::*;

/// How long one of Everything's own messages may take before it is given up on.
///
/// Short on purpose: the answer is the whole of a hover, so a message that has not come back by then
/// is a list that is busy or blocked — an elevated Everything answers nothing at all — and waiting
/// longer only holds the tick.
const ANSWER_MS: u32 = 250;

/// The scratch page every struct is written into and read back out of, in Everything's own process.
///
/// One page rather than an allocation per struct, because a row is three or four messages: the page
/// is made once for the list and every message writes over the front of it.
const PAGE_BYTES: usize = 4096;

/// Where a cell's text is written, in `u16`s from the front of that page: past every struct written
/// there, and far enough in that the two can never meet (`LVITEMW` is 88 bytes).
const TEXT_OFFSET: usize = 512;

/// How many characters of one cell are read, which is more than any file name and enough for any path
/// Everything draws in a column.
const TEXT_CHARS: usize = 512;

/// The most columns one row is read for.
///
/// **Every cell read is bounded by the list's own column count**, because Everything repeats the last
/// valid cell's text for a subitem past the columns it draws — a size or a date read as a path is
/// exactly what that answers — and this is the ceiling on a frame that reports more columns than it
/// has.
const MAX_COLUMNS: i32 = 16;

/// How long a page that could not be made is left alone before it is asked for again.
///
/// What makes it fail fails for the life of the run rather than once — Everything started elevated
/// cannot be opened by a process that is not, and `allow_lvm=0` leaves the list answering nothing —
/// so the ask is slow rather than repeated: no probe rate of its own against a process that will
/// never answer.
const DENIED_RETRY_MS: u64 = 10_000;

/// How much of a list view's client area a result list has to have before it is read as the rows.
///
/// The frame holds windows of the same class that are never shown a listing (Everything's own result
/// list focus stand-in), and a list with no room to draw a row in is not the one the user is pointing
/// at.
const THE_LIST_EDGE_PX: i32 = 50;

/// Everything's own process, and the one page in it that its list is read through.
///
/// **Everything reads the struct a message carries out of this process's own memory**, so a pointer
/// into the app would be answered against memory Everything cannot read: the page is what makes a
/// message about a row into the row's box. It is freed — in Everything's process, and with the
/// handle that opened it — when the reads it belongs to are dropped (see `EverythingReads`).
struct EverythingPage {
    process: HANDLE,
    remote: *mut core::ffi::c_void,
}

impl Drop for EverythingPage {
    fn drop(&mut self) {
        // Safety: the handle and the page are this struct's own, made by `of` and freed once here.
        unsafe {
            let _ = VirtualFreeEx(self.process, self.remote, 0, MEM_RELEASE);
            let _ = CloseHandle(self.process);
        }
    }
}

impl EverythingPage {
    /// A page in the process the list belongs to, or nothing where that process will not have one.
    ///
    /// Everything started elevated is the case that answers nothing: `OpenProcess` is refused for a
    /// process of a higher integrity level, and every message sent to it afterwards would be blocked
    /// as well. Neither handle nor page is asked for again here (see `DENIED_RETRY_MS`).
    fn of(list: HWND) -> Option<Self> {
        let mut pid = 0u32;
        if unsafe { GetWindowThreadProcessId(list, Some(&mut pid)) } == 0 || pid == 0 {
            return None;
        }

        // Safety: the pid is the list's own, and both handles made from it are this struct's until it
        // is dropped.
        unsafe {
            let process = OpenProcess(
                PROCESS_VM_OPERATION
                    | PROCESS_VM_READ
                    | PROCESS_VM_WRITE
                    | PROCESS_QUERY_LIMITED_INFORMATION,
                false,
                pid,
            )
            .ok()?;

            let remote = VirtualAllocEx(
                process,
                None,
                PAGE_BYTES,
                MEM_COMMIT | MEM_RESERVE,
                PAGE_READWRITE,
            );
            if remote.is_null() {
                let _ = CloseHandle(process);
                return None;
            }

            Some(Self { process, remote })
        }
    }

    /// One read-only message to the list, and the number it answered with — or nothing where it
    /// timed out or was blocked.
    ///
    /// A message that answered nothing cannot be told from one that answered zero, and the two
    /// answers that could be zero are the count of a list with no rows and a hit test that named no
    /// row: neither is a row to read.
    fn number(&self, list: HWND, msg: u32, wparam: usize, lparam: isize) -> Option<isize> {
        let mut answer = 0usize;
        // Safety: a message Everything's own window procedure answers, sent with no buffer of
        // either side's — the answer is a number and is taken through the out-parameter.
        let sent = unsafe {
            SendMessageTimeoutW(
                list,
                msg,
                WPARAM(wparam),
                LPARAM(lparam),
                SMTO_BLOCK | SMTO_ABORTIFHUNG,
                ANSWER_MS,
                Some(&mut answer),
            )
        };

        (sent.0 != 0).then_some(answer as isize)
    }

    /// One read-only message whose argument is a struct: written into the page, sent with the page's
    /// address, and read back out of the page.
    ///
    /// That round trip is the whole of what makes a message about a row answerable at all — see the
    /// note at the top of this module.
    fn exchange<T: Copy>(&self, list: HWND, msg: u32, wparam: usize, value: &mut T) -> bool {
        let size = std::mem::size_of::<T>();
        // No struct here is anywhere near the text area, and one that were would be written over the
        // text of the cell read before it.
        if size == 0 || size > TEXT_OFFSET {
            return false;
        }

        // Safety: both copies are one struct long, the page is one whole allocation of Everything's
        // and is longer than the part of it written into, and neither call may write past the buffer
        // it was handed — the message's own answer is never longer than the struct it was given.
        unsafe {
            if WriteProcessMemory(
                self.process,
                self.remote.cast_const(),
                (value as *const T).cast(),
                size,
                None,
            )
            .is_err()
            {
                return false;
            }

            let sent = SendMessageTimeoutW(
                list,
                msg,
                WPARAM(wparam),
                LPARAM(self.remote as isize),
                SMTO_BLOCK | SMTO_ABORTIFHUNG,
                ANSWER_MS,
                None,
            );
            if sent.0 == 0 {
                return false;
            }

            ReadProcessMemory(
                self.process,
                self.remote.cast_const(),
                (value as *mut T).cast(),
                size,
                None,
            )
            .is_ok()
        }
    }

    /// How many rows the list holds and how many columns it draws each of them with.
    fn shape(&self, list: HWND) -> Option<(i32, i32)> {
        let count = self.number(list, LVM_GETITEMCOUNT, 0, 0)? as i32;
        if count <= 0 {
            return None;
        }

        // The header is a control of its own in Everything's frame, and its column count is what a
        // row's cells are bounded by. A list that will not name one is read for the name alone.
        let columns = self
            .number(list, LVM_GETHEADER, 0, 0)
            .and_then(|header| {
                self.number(
                    HWND(header as *mut core::ffi::c_void),
                    HDM_GETITEMCOUNT,
                    0,
                    0,
                )
            })
            .unwrap_or(1) as i32;

        Some((count, columns.clamp(1, MAX_COLUMNS)))
    }

    /// One row's box, in the list's own client coordinates.
    fn row_box(&self, list: HWND, index: i32) -> Option<RECT> {
        // `LVIR_BOUNDS` is written into the left edge and the box is written over the whole struct:
        // the message takes its argument as a value and answers into the same buffer.
        let mut bounds = RECT {
            left: LVIR_BOUNDS as i32,
            ..Default::default()
        };

        self.exchange(list, LVM_GETITEMRECT, index as usize, &mut bounds)
            .then_some(bounds)
    }

    /// What one cell of a row says.
    fn cell(&self, list: HWND, index: i32, subitem: i32) -> Option<String> {
        // The text is written where the struct points, so the address it points at is Everything's:
        // it is taken from the page this process has just been given, and the struct carrying it is
        // written into that same page.
        //
        // Safety: the page is one whole allocation and the text area is inside it — `TEXT_OFFSET`
        // and `TEXT_CHARS` together are a quarter of `PAGE_BYTES`.
        let text_at = unsafe { self.remote.cast::<u16>().add(TEXT_OFFSET) };

        let mut item = LVITEMW {
            mask: LVIF_TEXT,
            iItem: index,
            iSubItem: subitem,
            pszText: PWSTR(text_at),
            cchTextMax: TEXT_CHARS as i32,
            ..Default::default()
        };
        if !self.exchange(list, LVM_GETITEMTEXTW, index as usize, &mut item) {
            return None;
        }

        let mut text = vec![0u16; TEXT_CHARS];
        // Safety: the buffer is `TEXT_CHARS` long and exactly that many `u16`s are read out of the
        // text area of the page, which is that long.
        unsafe {
            ReadProcessMemory(
                self.process,
                text_at.cast_const().cast(),
                text.as_mut_ptr().cast(),
                TEXT_CHARS * 2,
                None,
            )
            .ok()?;
        }

        let len = text.iter().position(|ch| *ch == 0).unwrap_or(text.len());
        Some(String::from_utf16_lossy(&text[..len]))
    }

    /// One row, as everything the roads ask a row for.
    fn row(&self, list: HWND, index: i32) -> Option<EverythingRow> {
        let (count, columns) = self.shape(list)?;
        if !row_is_in_the_list(index, count) {
            return None;
        }

        let box_of_row = self.row_box(list, index)?;
        // The list reports a row's box in its own client coordinates and every box this app holds is
        // in the screen's: the pointer, the preview and the box a hover is held by are all in screen
        // pixels, so the row's own corner is where the list's client area begins on the display.
        let mut origin = POINT {
            x: box_of_row.left,
            y: box_of_row.top,
        };
        if !unsafe { ClientToScreen(list, &mut origin) }.as_bool() {
            return None;
        }

        // The name is the row: a list that will not say what its first cell says has no file to name,
        // and the columns beside it are only read where there is a name to read them beside.
        let mut cells = vec![self.cell(list, index, 0)?];
        for subitem in 1..columns {
            let Some(cell) = self.cell(list, index, subitem) else {
                continue;
            };
            cells.push(cell);
        }

        Some(EverythingRow {
            index,
            list: list.0 as isize,
            bounds: RECT {
                left: origin.x,
                top: origin.y,
                right: origin.x + (box_of_row.right - box_of_row.left),
                bottom: origin.y + (box_of_row.bottom - box_of_row.top),
            },
            columns: columns > 1,
            cells,
        })
    }

    /// The row a point on the screen is on, where the list names one whose own box holds the point.
    fn row_at(&self, list: HWND, point: POINT) -> Option<EverythingRow> {
        let (count, _) = self.shape(list)?;

        let mut client = point;
        if !unsafe { ScreenToClient(list, &mut client) }.as_bool() {
            return None;
        }

        // **A point outside the list's client area is on no row, whatever the hit test says.** The
        // hit test names row 0 for a point above the first row and the last row the window draws for
        // one below it, and both of those rows have boxes that reach there — what refuses them is the
        // client area, which is where the rows stop being drawn (see `row_held_by`).
        let mut area = RECT::default();
        unsafe { GetClientRect(list, &mut area) }.ok()?;
        if !holds_point(area, client) {
            return None;
        }

        let mut hit = LVHITTESTINFO {
            pt: client,
            ..Default::default()
        };
        if !self.exchange(list, LVM_HITTEST, 0, &mut hit) {
            return None;
        }

        let top = self.number(list, LVM_GETTOPINDEX, 0, 0).unwrap_or(0) as i32;
        let index = row_held_by(client, hit.iItem, top, count, |index| {
            self.row_box(list, index)
        })?;

        self.row(list, index)
    }

    /// The row the keyboard is on: the focused one, or — where the list has none focused — the first
    /// selected one, which is the item a listing shows as current when the focus is elsewhere.
    fn focused_row(&self, list: HWND) -> Option<EverythingRow> {
        let (count, _) = self.shape(list)?;

        let index = [LVNI_FOCUSED, LVNI_SELECTED]
            .into_iter()
            .find_map(|flags| {
                // `usize::MAX` is `-1` to the list: the search starts before the first row, which is what
                // "the focused row" is asked as.
                self.number(list, LVM_GETNEXTITEM, usize::MAX, flags as isize)
                    .map(|index| index as i32)
                    .filter(|index| row_is_in_the_list(*index, count))
            })?;

        self.row(list, index)
    }
}

/// What one Everything listing is read through, kept in the resolver for as long as it describes the
/// list being read.
///
/// It goes with the list rather than with the frame because that is what the page is a page *of*: a
/// second Everything instance is another process, its list is another handle, and the page in the
/// first one's address space says nothing about it. Dropped, it frees that page (see `EverythingPage`).
pub(super) struct EverythingReads {
    list: isize,
    /// Nothing where no page could be made, or none has been asked for yet — see `page`.
    page: Option<EverythingPage>,
    /// When a page that is not there may be asked for again, which is what keeps an Everything this
    /// app may not read from being asked every tick.
    retry_at: Option<Instant>,
}

impl EverythingReads {
    fn of(list: isize) -> Self {
        Self {
            list,
            page: None,
            retry_at: None,
        }
    }

    /// The page the list is read through, made where one can be and the last failure is not fresh.
    fn page(&mut self) -> Option<&EverythingPage> {
        if self.page.is_none() {
            if self
                .retry_at
                .is_some_and(|retry_at| Instant::now() < retry_at)
            {
                return None;
            }

            self.page = EverythingPage::of(HWND(self.list as *mut core::ffi::c_void));
            self.retry_at = self
                .page
                .is_none()
                .then(|| Instant::now() + Duration::from_millis(DENIED_RETRY_MS));
        }

        self.page.as_ref()
    }
}

/// One row of Everything's result list, as the list itself describes it.
pub(super) struct EverythingRow {
    /// Zero-based, as the list numbers its rows — everything this app holds is one-based (see
    /// `HoveredItem`).
    pub(super) index: i32,
    /// The list the row is drawn in, which is the window an item of it reports as its own.
    pub(super) list: isize,
    /// The row's box on the screen.
    pub(super) bounds: RECT,
    /// What the row's cells say, cell 0 — the name — first, and the folder the file is in among the
    /// rest (see `row_file`).
    pub(super) cells: Vec<String>,
    /// Whether the list draws more than one column, which is what says a row's text is written into
    /// columns as wide as the view rather than drawn as a box of its own (see
    /// [`HoveredItem::draws_columns`]).
    pub(super) columns: bool,
}

impl EverythingRow {
    /// The item every road answers with, built out of the row's own text.
    ///
    /// Three of its facts are the row's — the name it draws, its box, and the window it is drawn in —
    /// and the fourth is the region a preview is kept off: a row of a details list draws its text
    /// across the row and nothing outside it, so the row whole is what `Avoid Details` keeps clear
    /// and its left end is what `Avoid Filename` narrows to the width the name is drawn at (see
    /// [`HoveredItem::avoid_box`]).
    ///
    /// **It carries no accessible value**, because the list has none to give: the item of a
    /// `SysListView32` row reports no value at all, and the path it would carry is the file's own
    /// entry read (see `row_file`) — a file read rather than a list read, paid twice for one hover
    /// and paid again on a share. The position is the whole of a row's identity, and the row's cells
    /// are what turn it into a file (see `item_path`).
    pub(super) fn item(&self) -> HoveredItem {
        HoveredItem {
            index: Some(self.index + 1),
            name: self.name().to_string(),
            value: None,
            bounds: self.bounds,
            text: Some(ItemText {
                all: self.bounds,
                name: self.bounds,
                columns: self.columns,
            }),
            native_window: self.list,
        }
    }

    /// The name the row draws, which is its first cell.
    fn name(&self) -> &str {
        self.cells
            .first()
            .map(|cell| cell.trim())
            .unwrap_or_default()
    }
}

/// The Everything frame a window stands in, where the app previews in that listing.
///
/// It is the branch every road takes before it asks UI Automation anything, and it is taken on the
/// class of the frame alone: the ask is a `GetAncestor` and two class compares, which is what keeps
/// the answer out of a walk that must never be made in this window (see the note at the top of this
/// module).
pub(super) fn the_everything_frame_of(window: HWND) -> Option<isize> {
    let frame = crate::shell::window_classes::the_frame_of(window);
    crate::shell::window_classes::everything_class_matches(frame).then_some(frame.0 as isize)
}

/// The result list of an Everything frame: the visible `SysListView32` in it with room to draw a row.
///
/// **Found by class and visibility, and never by control id.** 1.5.0.1423b draws the results at 10020,
/// and the windows around it — 10041's hidden `EVERYTHING_RESULT_LIST_FOCUS`, 10009's `SysHeader32`,
/// the status bar, the toolbar's search box — are windows of the same frame that answer none of what
/// a listing answers. An id is one version's arrangement of the frame; the class is the control's.
fn the_list_of(frame: HWND) -> Option<HWND> {
    struct TheSearch {
        found: Option<HWND>,
    }

    unsafe extern "system" fn the_first_visible_list(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let search = &mut *(lparam.0 as *mut TheSearch);
        if search.found.is_some() {
            return BOOL(0);
        }

        let mut class_name = [0u16; 256];
        // Safety: the class name is this window's and is written into this callback's own buffer.
        let len = GetClassNameW(hwnd, &mut class_name);
        if len <= 0
            || !utf16_contains_ascii_ignore_case(&class_name[..len as usize], b"syslistview32")
        {
            return BOOL(1);
        }

        // Safety: asks the desktop about a window it enumerated, and dereferences nothing.
        if !IsWindowVisible(hwnd).as_bool() {
            return BOOL(1);
        }

        let mut area = RECT::default();
        // Safety: as above, into this callback's own buffer.
        if GetClientRect(hwnd, &mut area).is_err()
            || area.right - area.left < THE_LIST_EDGE_PX
            || area.bottom - area.top < THE_LIST_EDGE_PX
        {
            return BOOL(1);
        }

        search.found = Some(hwnd);
        BOOL(1)
    }

    let mut search = TheSearch { found: None };
    // Safety: the callback is a plain function of this module's, and it is handed a pointer to a
    // struct that lives on this frame until the enumeration is over.
    unsafe {
        let _ = EnumChildWindows(
            frame,
            Some(the_first_visible_list),
            LPARAM(&mut search as *mut TheSearch as isize),
        );
    }

    search.found
}

/// The page the list is read through, made for the list it describes.
fn the_reads_of(resolver: &mut ItemResolver, list: HWND) -> Option<&EverythingPage> {
    let list = list.0 as isize;
    if resolver
        .everything
        .as_ref()
        .is_none_or(|reads| reads.list != list)
    {
        resolver.everything = Some(EverythingReads::of(list));
    }

    resolver.everything.as_mut()?.page()
}

/// One row of the Everything list under `frame`, read by its index — zero-based, as the list numbers
/// its rows.
pub(super) fn row_of(
    resolver: &mut ItemResolver,
    frame: isize,
    index: i32,
) -> Option<EverythingRow> {
    let list = the_list_of(HWND(frame as *mut core::ffi::c_void))?;
    the_reads_of(resolver, list)?.row(list, index)
}

/// The row a point on the screen is on, in the Everything list under `frame`.
pub(super) fn row_at(
    resolver: &mut ItemResolver,
    frame: isize,
    point: POINT,
) -> Option<EverythingRow> {
    let list = the_list_of(HWND(frame as *mut core::ffi::c_void))?;
    the_reads_of(resolver, list)?.row_at(list, point)
}

/// The row the keyboard is on in the Everything list under `frame`.
pub(super) fn focused_row(resolver: &mut ItemResolver, frame: isize) -> Option<EverythingRow> {
    let list = the_list_of(HWND(frame as *mut core::ffi::c_void))?;
    the_reads_of(resolver, list)?.focused_row(list)
}

/// The file a row's cells name, or nothing where they name none.
///
/// **The row's own text is the only witness there is.** Everything holds no path for a row beyond the
/// columns it draws, and which columns those are is the user's arrangement: so the name is cell 0 and
/// the folder is the first other cell that *is* one — an absolute path that is a directory, which a
/// size (`51,964 KB`), a date or a count never is — and the join is gated the way every route here is
/// gated (`normalize_media_path` reads the file's own entry, so a path is only taken where the file is
/// there, is a file, and is a kind this build shows).
///
/// A folder, a file no kind previews, and a row whose columns show no folder all answer nothing.
/// Nothing is looked up by name in any folder: a search's rows come from every folder the user pointed
/// it at, and the same name in two of them is two files.
pub(super) fn row_file(cells: &[String]) -> Option<PathBuf> {
    let name = cells.first()?.trim();
    if name.is_empty() {
        return None;
    }

    let folder = cells.iter().skip(1).find_map(|cell| {
        let cell = cell.trim();
        let path = Path::new(cell);
        (!cell.is_empty() && path.is_absolute() && path.is_dir()).then_some(path)
    })?;

    normalize_media_path(folder.join(name))
}

/// The file the item `item_file_path` was handed stands for, in the Everything listing it was read in:
/// the row the item names, and the file that row's cells name (see `row_file`).
///
/// It is the one place an Everything item becomes a file, so the pointer's row, the keyboard's row and
/// any walk by index all answer through it.
pub(super) fn item_path(
    resolver: &mut ItemResolver,
    window: &ItemWindow,
    item: &HoveredItem,
) -> Option<PathBuf> {
    let answer =
        row_of(resolver, window.frame, item.index? - 1).and_then(|row| row_file(&row.cells));

    // The row the answer was read from against the file it named, because a row that names no file is
    // the ordinary case here — a folder, a kind nothing previews, a column set without the folder's
    // path — and reads here as `-` rather than as a run that never looked.
    note_pin_click!(
        "everything item  {}  -> {}",
        item.name,
        answer
            .as_deref()
            .map_or_else(|| "-".to_string(), |path| path.display().to_string()),
    );

    answer
}

/// Whether an index the list answered is one of its rows.
///
/// A focus, a hit test's answer and a top index are numbers that a list which moved under the answer
/// can leave outside it — an emptied results list answers a focus of `-1` — and an index outside the
/// list is no row at all.
pub(super) fn row_is_in_the_list(index: i32, count: i32) -> bool {
    index >= 0 && index < count
}

/// The rows a hit test's answer may name, in the order they are looked at: the index as it came, and
/// the same index read from the first row the list draws.
///
/// **Everything's hit test names a row whether or not the point is on one** — row 0 for any point
/// above the first row, and the last row the window draws for a point below it (both measured) — and
/// its flags say nothing either: a point well inside a wide row answers `LVHT_NOWHERE`. So its answer
/// is only taken where the row's own box holds the point, which is what `row_held_by` asks. The
/// second candidate is for a list that numbers its answer from the row it has scrolled to rather than
/// from the first row of the list, which a results list does far more often than a folder view does.
pub(super) fn hit_rows(hit: i32, top: i32, count: i32) -> [Option<i32>; 2] {
    [hit, hit + top].map(|index| row_is_in_the_list(index, count).then_some(index))
}

/// The row a point is on, out of what a hit test answered and the boxes the list draws: the first of
/// [`hit_rows`] whose own box holds the point, and nothing where none of them does.
pub(super) fn row_held_by(
    point: POINT,
    hit: i32,
    top: i32,
    count: i32,
    box_of_row: impl Fn(i32) -> Option<RECT>,
) -> Option<i32> {
    hit_rows(hit, top, count)
        .into_iter()
        .flatten()
        .find(|index| box_of_row(*index).is_some_and(|box_of_row| holds_point(box_of_row, point)))
}

/// Whether a box holds a point, the way every box here is read: the right and bottom edges are
/// outside it, which is what makes two rows drawn edge to edge hold every point between them once.
pub(super) fn holds_point(rect: RECT, point: POINT) -> bool {
    point.x >= rect.left && point.x < rect.right && point.y >= rect.top && point.y < rect.bottom
}
