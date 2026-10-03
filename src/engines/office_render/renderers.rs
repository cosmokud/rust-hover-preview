use crate::config::config::{image_decode_limits, read_within_budget};
use crate::engines::document_cache::{self, PageKind};
use crate::formats::office_formats::OfficeApp;
use crate::paths::plain_path;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU32, Ordering};
use std::time::Duration;

use windows::core::VARIANT;
use windows::Win32::Foundation::HGLOBAL;
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, GetClipboardData, OpenClipboard,
};
use windows::Win32::System::Memory::{GlobalLock, GlobalSize, GlobalUnlock};
use windows::Win32::System::Ole::CF_DIB;

use super::com::{path_variant, Object};

/// The width a slide is exported at, bounded: enough for any preview box, and
/// never a poster.
const MIN_SLIDE_EXPORT_WIDTH: u32 = 640;
pub(super) const MAX_SLIDE_EXPORT_WIDTH: u32 = 1920;
/// Paths at least this long are handed to Office as a copy in the temp folder.
/// Office is not a long-path consumer of the plain form the app converts to.
const MAX_OFFICE_PATH: usize = 240;
/// How much of a worksheet's used range is copied out when the machine has no
/// printer to export a page with: the corner a person sees first, not every row
/// the sheet holds.
const PICTURE_MAX_ROWS: i32 = 40;
const PICTURE_MAX_COLUMNS: i32 = 14;
/// The most a worksheet's picture may be, in pixels.
///
/// What a preview can show is bounded by the display, and a report whose rows are
/// tall — wrapped headings, merged cells — or whose columns are wide would
/// otherwise be copied out at several million pixels: slow to copy, slow to write
/// and slow to draw, all to show a corner no preview can hold. The window the
/// picture is taken from is cut down until what it would draw fits these — and they
/// are deliberately smaller than a display, because a picture is shown at its own
/// size (the configured scale) rather than fitted to the screen the way a page is,
/// and a corner of a sheet filling the screen is not a preview of anything.
const PICTURE_MAX_PIXELS_WIDTH: f64 = 900.0;
const PICTURE_MAX_PIXELS_HEIGHT: f64 = 700.0;
/// The least of a sheet still worth copying: less than this shows too little to be
/// a preview of anything.
const PICTURE_MIN_ROWS: i32 = 4;
const PICTURE_MIN_COLUMNS: i32 = 3;
/// How many times the window may be cut down before it is copied as it stands.
const PICTURE_FIT_ATTEMPTS: usize = 6;
/// Points to pixels, as Excel lays a sheet out at 96 DPI: a point is a 72nd of an
/// inch.
const PICTURE_PIXELS_PER_POINT: f64 = 96.0 / 72.0;
/// The clipboard is shared with every other process: a look that cannot open it,
/// or finds nothing in it, is retried this many times.
const CLIPBOARD_ATTEMPTS: usize = 5;
/// How many times a picture is asked for, and how long between the asks.
///
/// The first copy of a workbook Excel has just opened is regularly refused — an
/// exception saying the `CopyPicture` property of the range could not be got, which
/// is Excel still laying the sheet out rather than anything about the document — and
/// the next attempt a moment later is answered. Asking only once is what left whole
/// workbooks with no preview while their neighbours in the same folder had one.
const PICTURE_COPY_ATTEMPTS: usize = 3;
const PICTURE_COPY_RETRY_MS: u64 = 120;
/// `xlScreen` and `xlBitmap`: the appearance and format `CopyPicture` is asked for.
const XL_SCREEN: i32 = 1;
const XL_BITMAP: i32 = 2;
/// How much of a worksheet a page is exported from: the top-left window of its used
/// range, in cells.
///
/// A printed page holds fewer than a hundred rows and a few dozen columns at any
/// paper size and scale, so this is several times a page and page 1 is always inside
/// it. What it is not is a hundred thousand rows, which is what the used range of a
/// workbook of a few kilobytes can be: formatting that runs down a column makes a
/// used range out of cells that hold nothing, and `ExportAsFixedFormat` lays out
/// every one of them to find out where page 1 ends.
const PAGE_MAX_ROWS: i32 = 128;
const PAGE_MAX_COLUMNS: i32 = 64;

/// What suppresses a dialog in each application: Word and PowerPoint take a
/// level, Excel a boolean.
pub(super) fn alerts_off(app_kind: OfficeApp) -> VARIANT {
    match app_kind {
        OfficeApp::Word => VARIANT::from(0i32),       // wdAlertsNone
        OfficeApp::Excel => VARIANT::from(false),     // DisplayAlerts is a boolean
        OfficeApp::PowerPoint => VARIANT::from(1i32), // ppAlertsNone
    }
}

// --------------------------------------------------------------- the render

pub(super) fn render_word(app: &Object, source: &Path, target: &RenderTarget) -> bool {
    let Some(documents) = app.member("Documents") else {
        return false;
    };
    let opened = documents.call(
        "Open",
        &[
            ("FileName", path_variant(source)),
            ("ReadOnly", VARIANT::from(true)),
            ("AddToRecentFiles", VARIANT::from(false)),
            ("ConfirmConversions", VARIANT::from(false)),
            // The document opens in a hidden window even when the application
            // itself is the user's and visible.
            ("Visible", VARIANT::from(false)),
        ],
    );
    let Some(document) = opened.and_then(Object::from_variant) else {
        return false;
    };

    let rendered = document
        .call(
            "ExportAsFixedFormat",
            &[
                ("OutputFileName", path_variant(&target.file("pdf"))),
                ("ExportFormat", VARIANT::from(17i32)), // wdExportFormatPDF
                ("OpenAfterExport", VARIANT::from(false)),
                ("Range", VARIANT::from(3i32)), // wdExportFromTo
                ("From", VARIANT::from(1i32)),
                ("To", VARIANT::from(1i32)),
            ],
        )
        .is_some();

    let _ = document.call("Close", &[("SaveChanges", VARIANT::from(false))]);
    rendered
}

pub(super) fn render_excel(
    app: &Object,
    source: &Path,
    target: &RenderTarget,
    retrying: bool,
) -> bool {
    let Some(workbooks) = app.member("Workbooks") else {
        return false;
    };
    let opened = workbooks.call(
        "Open",
        &[
            ("FileName", path_variant(source)),
            ("UpdateLinks", VARIANT::from(0i32)),
            ("ReadOnly", VARIANT::from(true)),
            ("AddToMru", VARIANT::from(false)),
            ("IgnoreReadOnlyRecommended", VARIANT::from(true)),
        ],
    );
    let Some(workbook) = opened.and_then(Object::from_variant) else {
        return false;
    };

    // A workbook's first page is what it prints, and exporting one goes through
    // the print pipeline: Excel needs a printer on the machine for it, and a
    // machine with none — no printer at all, which is not the same as a sheet
    // without a print area — cannot export a page however it is asked.
    //
    // The picture is the no-printer answer. On a machine that has a printer an
    // export that came to nothing is the instance declining — the state a fresh
    // instance is about to be asked in — rather than the document failing, so the
    // picture is left to an attempt that is already a retry, where it is the last
    // thing tried before the document is given up on. Copying a picture from an
    // instance that has just refused to export is a detour of its own: the copy
    // is refused by the same instance, after the retries its clipboard needs.
    let has_printer = printer_installed(app);
    let printed = has_printer && export_first_page(&workbook, target);
    let rendered =
        printed || ((!has_printer || retrying) && copy_used_range_picture(&workbook, target));

    let _ = workbook.call("Close", &[("SaveChanges", VARIANT::from(false))]);
    rendered
}

/// The first worksheet's first printed page, as a PDF.
///
/// A used range that reaches past a page is exported as the top-left window of
/// itself rather than as the worksheet. `ExportAsFixedFormat` has to lay a sheet out
/// for printing to find out what page 1 is, and that walk is over the used range
/// rather than over the data in it — so a workbook whose formatting runs down a
/// column pays for a hundred thousand empty cells to draw a page that holds forty of
/// them. A window is enough for the page that is asked for, because pagination runs
/// from the top-left: page 1 of the window is page 1 of the sheet. A used range that
/// already fits the window is left alone, so an ordinary sheet is exported exactly
/// as it was before there was one.
fn export_first_page(workbook: &Object, target: &RenderTarget) -> bool {
    let Some(sheet) = workbook
        .member("Worksheets")
        .and_then(|sheets| sheets.item(1))
    else {
        return false;
    };
    let output = target.file("pdf");

    if let Some(window) = used_window(&sheet) {
        if export_page(&window, &output) {
            return true;
        }
        // Nothing usable came of it, so whatever it left is removed before the
        // sheet is asked: a file still lying there would answer for the attempt
        // that comes after it.
        let _ = std::fs::remove_file(&output);
    }

    export_page(&sheet, &output)
}

/// One page of `source` — a worksheet, or the window of one — written to `output`.
///
/// The page is asked for by number, and `1` to `1` is the first one. What the page
/// looks like is the sheet's own print setup either way, so a window is not a
/// different page from the sheet's; it is the same page, reached without laying out
/// everything below it.
fn export_page(source: &Object, output: &Path) -> bool {
    let exported = source
        .call(
            "ExportAsFixedFormat",
            &[
                ("Type", VARIANT::from(0i32)), // xlTypePDF
                ("Filename", path_variant(output)),
                ("From", VARIANT::from(1i32)),
                ("To", VARIANT::from(1i32)),
                ("OpenAfterPublish", VARIANT::from(false)),
            ],
        )
        .is_some();

    exported && output.exists()
}

/// The top-left window of a worksheet's used range, when the range reaches past it.
///
/// `None` when it already fits — there is nothing to bound, and the sheet answers
/// for itself, print areas and all — and when it cannot be measured at all, which
/// leaves the sheet to be exported as it is rather than refusing it.
fn used_window(sheet: &Object) -> Option<Object> {
    let used = sheet.member("UsedRange")?;
    let rows = collection_count(used.member("Rows"))?;
    let columns = collection_count(used.member("Columns"))?;

    if rows <= PAGE_MAX_ROWS && columns <= PAGE_MAX_COLUMNS {
        return None;
    }

    resize_range(
        &used,
        rows.min(PAGE_MAX_ROWS),
        columns.min(PAGE_MAX_COLUMNS),
    )
}

/// Ask the application something that only an application with a workbook open will
/// answer, and put it back as it was found.
///
/// Excel refuses `Application.Calculation` while nothing is open — a read answers
/// nothing at all, and a write raises "Unable to set the Calculation property of the
/// Application class", which is what the engine's first attempt at it was met with,
/// made as it was before any document existed. The setting is wanted *before* a
/// document is opened, because what it decides is whether opening one recalculates
/// it, so a scratch workbook is opened for the asking and closed straight after.
///
/// It is added only when there is nothing open, so an instance the user is working
/// in is not handed a stray `Book1`, and it is closed rather than kept: what is
/// wanted is a setting, not a document.
///
/// Only the probe below the tests uses it now. The engine used to hold Excel in
/// manual calculation so that a page would be the one the file was saved as, and it
/// is not worth doing: a workbook is worked out again as it is opened whatever the
/// mode says — `excel_calculation_probe` measures exactly that — so the setting buys
/// nothing and costs a workbook to make.
#[cfg(test)]
pub(super) fn with_a_workbook<T>(app: &Object, ask: impl FnOnce() -> T) -> T {
    let scratch = (collection_count(app.member("Workbooks")) == Some(0))
        .then(|| app.member("Workbooks"))
        .flatten()
        .and_then(|books| books.call("Add", &[]))
        .and_then(Object::from_variant);

    let answer = ask();

    if let Some(book) = scratch {
        let _ = book.call("Close", &[("SaveChanges", VARIANT::from(false))]);
    }

    answer
}

/// Whether the machine has a printer, which is what an export to a page needs.
///
/// Excel answers `ActivePrinter` with a name when there is one, and with the
/// sentence "unknown printer (check your Control Panel)" when there is not — so
/// the question is asked before the export rather than answered by its failure.
fn printer_installed(app: &Object) -> bool {
    app.value("ActivePrinter")
        .map(|value| value.to_string())
        .map(|name| !name.trim().is_empty() && !name.contains("unknown printer"))
        .unwrap_or(false)
}

/// The used range's top-left, copied out of Excel as a picture.
///
/// This is what a machine with no printer gets instead of a page: the range is
/// copied the way a person copies it — `CopyPicture` — and the bitmap Excel puts
/// on the clipboard is decoded and written out as a PNG. It was a BMP until this
/// — the one format that is exactly the bytes the clipboard holds — and what that
/// cost was a page of five to forty megabytes for a picture of flat-coloured
/// cells: written to the cache, read back out of it, and given up and written
/// again on every hover that lost it, where the same picture as a PNG is several
/// times smaller and no different (see `write_png`). What it shows is the corner
/// of the sheet a person would see first rather than the sheet's printed layout,
/// which is the most such a machine can produce.
fn copy_used_range_picture(workbook: &Object, target: &RenderTarget) -> bool {
    let Some(sheet) = workbook
        .member("Worksheets")
        .and_then(|sheets| sheets.item(1))
    else {
        return false;
    };
    let Some(used) = sheet.member("UsedRange") else {
        return false;
    };
    let (Some(rows), Some(columns)) = (
        collection_count(used.member("Rows")),
        collection_count(used.member("Columns")),
    ) else {
        return false;
    };

    let Some(range) = picture_range(&used, rows, columns) else {
        return false;
    };

    for attempt in 0..PICTURE_COPY_ATTEMPTS {
        let copied = range
            .call_args(
                "CopyPicture",
                &[VARIANT::from(XL_SCREEN), VARIANT::from(XL_BITMAP)],
            )
            .is_some();
        if copied {
            if let Some(dib) = clipboard_dib() {
                return write_png(&target.file("png"), &dib).is_ok();
            }
        }

        if attempt + 1 < PICTURE_COPY_ATTEMPTS {
            std::thread::sleep(Duration::from_millis(PICTURE_COPY_RETRY_MS));
        }
    }

    false
}

/// The window of the used range the picture is taken from: its top-left corner, cut
/// down until what it would draw fits the box a preview could ever show.
///
/// Only the range's own measurements are asked for — never its cells — so a sheet
/// of a million rows costs the same as a small one, and what comes back is the
/// corner a person would see first rather than a page of it. Each cut is in
/// proportion to how far over the budget the range is, so a sheet whose rows are
/// tall keeps as many of them as the budget allows instead of being halved until a
/// sliver is left.
fn picture_range(used: &Object, rows: i32, columns: i32) -> Option<Object> {
    let mut rows = rows.clamp(1, PICTURE_MAX_ROWS);
    let mut columns = columns.clamp(1, PICTURE_MAX_COLUMNS);
    let mut range = resize_range(used, rows, columns);

    for _ in 0..PICTURE_FIT_ATTEMPTS {
        let Some(current) = range.as_ref() else {
            break;
        };
        let (Some(width), Some(height)) =
            (point_size(current, "Width"), point_size(current, "Height"))
        else {
            break;
        };

        let width_px = width * PICTURE_PIXELS_PER_POINT;
        let height_px = height * PICTURE_PIXELS_PER_POINT;
        if width_px <= PICTURE_MAX_PIXELS_WIDTH && height_px <= PICTURE_MAX_PIXELS_HEIGHT {
            break;
        }

        let fitted_rows = if height_px > PICTURE_MAX_PIXELS_HEIGHT {
            (rows as f64 * PICTURE_MAX_PIXELS_HEIGHT / height_px).floor() as i32
        } else {
            rows
        };
        let fitted_columns = if width_px > PICTURE_MAX_PIXELS_WIDTH {
            (columns as f64 * PICTURE_MAX_PIXELS_WIDTH / width_px).floor() as i32
        } else {
            columns
        };

        let fitted_rows = fitted_rows.clamp(PICTURE_MIN_ROWS, rows);
        let fitted_columns = fitted_columns.clamp(PICTURE_MIN_COLUMNS, columns);
        // Whatever is left is smaller than a cell: copy it as it stands.
        if fitted_rows == rows && fitted_columns == columns {
            break;
        }

        rows = fitted_rows;
        columns = fitted_columns;
        range = resize_range(used, rows, columns);
    }

    range
}

/// The top-left window of a range, `rows` by `columns` cells of it.
fn resize_range(range: &Object, rows: i32, columns: i32) -> Option<Object> {
    range
        .call_args("Resize", &[VARIANT::from(rows), VARIANT::from(columns)])
        .and_then(Object::from_variant)
}

/// One of a range's own measurements, in points.
fn point_size(range: &Object, property: &str) -> Option<f64> {
    range
        .value(property)
        .and_then(|value| f64::try_from(&value).ok())
}

/// How many items a collection holds.
fn collection_count(collection: Option<Object>) -> Option<i32> {
    collection
        .and_then(|collection| collection.value("Count"))
        .and_then(|count| i32::try_from(&count).ok())
}

/// The bitmap Excel has just put on the clipboard.
fn clipboard_dib() -> Option<Vec<u8>> {
    clipboard_dib_inner(true)
}

/// `empty` says whether the clipboard is cleared once the picture has been taken.
/// It is, in the app: leaving Excel's copy on the clipboard is what makes Excel ask,
/// on its way out, whether a large amount of information should stay there.
fn clipboard_dib_inner(empty: bool) -> Option<Vec<u8>> {
    for attempt in 0..CLIPBOARD_ATTEMPTS {
        if let Some(dib) = read_clipboard_dib(empty) {
            return Some(dib);
        }
        if attempt + 1 < CLIPBOARD_ATTEMPTS {
            std::thread::sleep(Duration::from_millis(30));
        }
    }

    None
}

fn read_clipboard_dib(empty: bool) -> Option<Vec<u8>> {
    unsafe {
        if OpenClipboard(None).is_err() {
            return None;
        }

        let mut dib = None;
        if let Ok(handle) = GetClipboardData(CF_DIB.0 as u32) {
            let handle = HGLOBAL(handle.0);
            let size = GlobalSize(handle);
            let pointer = GlobalLock(handle) as *const u8;
            if !pointer.is_null() {
                if size > 0 {
                    dib = Some(std::slice::from_raw_parts(pointer, size).to_vec());
                }
                let _ = GlobalUnlock(handle);
            }
        }

        // The picture is taken rather than borrowed. Leaving it on the clipboard is
        // what makes Excel ask, on its way out, whether a large amount of
        // information should stay there — a dialog no one is present to answer,
        // which holds the quit and with it the worker.
        if dib.is_some() && empty {
            let _ = EmptyClipboard();
        }

        let _ = CloseClipboard();
        dib
    }
}

/// What the clipboard held, decoded and written as a PNG file.
///
/// A DIB is a `BITMAPINFO` and its pixels, and a BMP file is those bytes with a
/// fourteen-byte header in front of them — which is what the decoder is handed, a
/// headerless DIB being a file no decoder opens. What comes out is written as a PNG
/// because of what the picture is: a screenshot of flat-coloured cells, which is the
/// case that format is best at, so the same picture costs the page cache and the
/// hover that reads it back a fraction of what the bitmap did.
///
/// It is written beside its name and moved into place, so that a render which is
/// ended part-way through the write leaves something that is obviously not a
/// picture rather than half of one under the name a page is read from.
fn write_png(path: &Path, dib: &[u8]) -> std::io::Result<()> {
    let offset = dib_pixel_offset(dib).unwrap_or(54);
    let mut file = Vec::with_capacity(dib.len() + 14);
    file.extend_from_slice(b"BM");
    file.extend_from_slice(&((dib.len() + 14) as u32).to_le_bytes());
    file.extend_from_slice(&0u16.to_le_bytes()); // reserved
    file.extend_from_slice(&0u16.to_le_bytes()); // reserved
    file.extend_from_slice(&offset.to_le_bytes());
    file.extend_from_slice(dib);

    // Read under the budget every other decode of this app is read under: the picture is
    // Excel's, but the bytes reach this side through the clipboard, which is not this app's.
    let mut reader = image::ImageReader::new(std::io::Cursor::new(&file)).with_guessed_format()?;
    reader.limits(image_decode_limits());
    let picture = reader.decode().map_err(picture_error)?;

    let mut png = Vec::new();
    picture
        .write_to(&mut std::io::Cursor::new(&mut png), image::ImageFormat::Png)
        .map_err(picture_error)?;

    let writing = path.with_extension("png.writing");
    std::fs::write(&writing, &png)?;
    std::fs::rename(&writing, path)
}

/// A decode or an encode that failed, as the error the write answers with: the caller is a
/// render, and a picture it could not write is answered the way one it could not copy is —
/// with no page.
fn picture_error(error: image::ImageError) -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidData, error)
}

/// Where a DIB's pixels start, past its header and its colour table.
fn dib_pixel_offset(dib: &[u8]) -> Option<u32> {
    let header = u32::from_le_bytes(dib.get(0..4)?.try_into().ok()?) as usize;
    if header < 40 || header > dib.len() {
        return None;
    }

    let bit_count = u16::from_le_bytes(dib.get(14..16)?.try_into().ok()?) as u32;
    let used_colors = u32::from_le_bytes(dib.get(32..36)?.try_into().ok()?);
    let palette_entries = if bit_count <= 8 {
        if used_colors != 0 {
            used_colors
        } else {
            1u32 << bit_count
        }
    } else {
        0
    };

    Some((14 + header) as u32 + palette_entries * 4)
}

pub(super) fn render_powerpoint(
    app: &Object,
    source: &Path,
    target: &RenderTarget,
    width: u32,
    height: u32,
) -> bool {
    let Some(presentations) = app.member("Presentations") else {
        return false;
    };
    let opened = presentations.call(
        "Open",
        &[
            ("FileName", path_variant(source)),
            ("ReadOnly", VARIANT::from(-1i32)), // msoTrue
            ("Untitled", VARIANT::from(0i32)),  // msoFalse
            ("WithWindow", VARIANT::from(0i32)),
        ],
    );
    let Some(presentation) = opened.and_then(Object::from_variant) else {
        return false;
    };

    // A slide is exported as an image rather than the deck as a PDF: the export
    // writes slide 1 alone instead of every slide the deck holds. It is reached
    // through the collection's item, which PowerPoint exposes as a method —
    // asking `Slides` itself for one is answered with "member not found".
    let rendered = presentation
        .member("Slides")
        .and_then(|slides| slides.item(1))
        .map(|slide| {
            let (export_width, export_height) = slide_export_size(&presentation, width, height);
            slide
                .call(
                    "Export",
                    &[
                        ("FileName", path_variant(&target.file("png"))),
                        ("FilterName", VARIANT::from("PNG")),
                        ("ScaleWidth", VARIANT::from(export_width)),
                        ("ScaleHeight", VARIANT::from(export_height)),
                    ],
                )
                .is_some()
        })
        .unwrap_or(false);

    let _ = presentation.call("Close", &[]);
    rendered
}

/// A single-precision property. PowerPoint records a slide's size as one, and
/// automation may hand a number back as either width.
fn single_of(value: &VARIANT) -> Option<f32> {
    f64::try_from(value).ok().map(|value| value as f32)
}

/// The width a slide is exported at for a preview that asked for `width`: what the
/// box the preview may fill leaves, bounded so that a display far larger than any
/// slide is still shown a page rather than a poster — and so that the export never
/// grows with the display past this, whatever the room is.
pub(super) fn slide_export_width(width: u32) -> u32 {
    width.clamp(MIN_SLIDE_EXPORT_WIDTH, MAX_SLIDE_EXPORT_WIDTH)
}

/// The size slide 1 is exported at: the width the preview asked for, bounded,
/// and the height that keeps the slide's own aspect ratio.
fn slide_export_size(presentation: &Object, width: u32, height: u32) -> (i32, i32) {
    let export_width = slide_export_width(width) as i32;

    let slide_width = presentation
        .member("PageSetup")
        .and_then(|setup| setup.value("SlideWidth"))
        .and_then(|value| single_of(&value))
        .filter(|value| *value > 1.0);
    let slide_height = presentation
        .member("PageSetup")
        .and_then(|setup| setup.value("SlideHeight"))
        .and_then(|value| single_of(&value))
        .filter(|value| *value > 1.0);

    let ratio = match (slide_width, slide_height) {
        (Some(slide_width), Some(slide_height)) => slide_height / slide_width,
        _ => height.max(1) as f32 / width.max(1) as f32,
    };

    (
        export_width,
        ((export_width as f32 * ratio).round() as i32).max(1),
    )
}

// -------------------------------------------------------------- the sources

/// The document handed to the engine, and whether it is a copy made for it.
///
/// Office is not a consumer of the verbatim `\\?\` paths the rest of this app
/// canonicalizes to, a path longer than the plain limit is one it cannot open at
/// all, and a document carrying a zone identifier is one Word and Excel open in
/// Protected View — where the export is refused. Any of those is answered with a
/// copy in the temp folder, which also keeps the engine from locking a file the
/// user may be working in.
pub(super) struct PreparedSource {
    pub(super) path: PathBuf,
    copy: bool,
}

impl PreparedSource {
    pub(super) fn new(source: &Path) -> Self {
        let plain = plain_path(source);

        if plain.chars().count() < MAX_OFFICE_PATH && !has_zone_identifier(&plain) {
            return Self {
                path: PathBuf::from(plain),
                copy: false,
            };
        }

        match copy_to_temp(&plain) {
            Some(path) => Self { path, copy: true },
            None => Self {
                path: PathBuf::from(plain),
                copy: false,
            },
        }
    }

    pub(super) fn cleanup(&self) {
        if self.copy {
            let _ = std::fs::remove_file(&self.path);
        }
    }
}

pub(super) fn has_zone_identifier(path: &str) -> bool {
    std::fs::metadata(format!("{path}:Zone.Identifier")).is_ok()
}

fn copy_to_temp(source: &str) -> Option<PathBuf> {
    static COUNTER: AtomicU32 = AtomicU32::new(0);

    let folder = scratch_folder();
    std::fs::create_dir_all(&folder).ok()?;

    let extension = Path::new(source)
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or("bin");
    let name = format!(
        "render-{}-{}.{extension}",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    );
    let target = folder.join(name);

    std::fs::copy(source, &target).ok()?;
    Some(target)
}

// ---------------------------------------------------------- the scratch file

/// Where a render is written before it is read back out of it.
///
/// Office's export calls take a file name rather than a stream, so a page has to land somewhere
/// before it can be kept. It is this app's own folder under the temp folder — the same one a
/// document Office will not open is copied into, and the folder the kept pages sit inside — and
/// the file is deleted the moment it has been read, so what is on disk is one render in flight
/// and nothing else.
fn scratch_folder() -> PathBuf {
    document_cache::temp_folder()
}

/// Where one render writes its page. Which file it is — which extension — is the
/// renderer's to choose, since what a document can be drawn from is not known
/// until it has been asked.
pub(super) struct RenderTarget {
    folder: PathBuf,
    stem: String,
}

impl RenderTarget {
    pub(super) fn file(&self, extension: &str) -> PathBuf {
        self.folder.join(format!("{}.{extension}", self.stem))
    }
}

/// A name of this render's own, so that a file left behind by a process which was
/// ended mid-render is never mistaken for the page a later one just wrote.
pub(super) fn render_target() -> Option<RenderTarget> {
    static COUNTER: AtomicU32 = AtomicU32::new(0);

    let folder = scratch_folder();
    std::fs::create_dir_all(&folder).ok()?;

    Some(RenderTarget {
        folder,
        stem: format!(
            "page-{}-{}",
            std::process::id(),
            COUNTER.fetch_add(1, Ordering::Relaxed)
        ),
    })
}

/// Read the page a render wrote and delete it, whichever of the files a render can
/// produce it turned out to be.
///
/// Every candidate is removed whether or not it is the one read — an empty file
/// Office gave up on, or one this app's own reading could not take, is not left for
/// the next attempt to find. What comes back is the bytes and the kind they are.
pub(super) fn take_render(target: &RenderTarget) -> Option<(PageKind, Vec<u8>)> {
    let mut taken = None;

    // The kinds a render writes: a page exported as a PDF, a slide exported as a PNG and a
    // workbook's picture, which is a PNG since the picture stopped being written as a bitmap
    // (see `write_png`). A `.bmp` is not written by anything here any more — one left in the
    // folder by a run before this one is one this app's start clears away with the rest of what
    // a run leaves in the temp folder (`document_cache::discard_leftovers`).
    for kind in [PageKind::Pdf, PageKind::Png] {
        let path = target.file(kind.extension());
        if std::fs::metadata(&path).is_err() {
            continue;
        }

        // The export is read under the budget every other file read is under: a page
        // this app asked for is a page-sized file, and a document whose export came out
        // many times that is answered as a render that produced nothing.
        let bytes = read_within_budget(&path);
        let _ = std::fs::remove_file(&path);

        if taken.is_none() {
            if let Some(bytes) = bytes {
                if !bytes.is_empty() {
                    taken = Some((kind, bytes));
                }
            }
        }
    }

    // A picture is written beside its name and moved into place, so a process that
    // was ended inside that write leaves the half it had written under this name.
    let _ = std::fs::remove_file(target.file("png.writing"));

    taken
}
