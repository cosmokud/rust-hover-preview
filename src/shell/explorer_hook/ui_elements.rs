//! The UI Automation reads: the names that mean a container rather than a file, the
//! walk from the point up to the item under it, and the boxes that item's text is
//! drawn in.
//!
//! Everything here is one batched request per element where the API allows it,
//! because the cost of this path is crossings into the shell and the properties are
//! asked together for that reason rather than one call at a time.

use super::*;

/// Names that indicate container elements, not actual files
pub(super) const CONTAINER_NAMES: &[&str] = &[
    "Items View",
    "Folder View",
    "Shell Folder View",
    "ShellView",
    "UIItemsView",
    "DirectUIHWND",
    "Search Results",
    "File list",
    "Name",
    "Date modified",
    "Type",
    "Size",
    "Date",
    "Date created",
    "Details",
    "List",
    "Content",
    "Tiles",
    "Large icons",
    "Medium icons",
    "Small icons",
    "Extra large icons",
    "Item",
    "Group",
    "Header",
];

/// Patterns that suggest a value might be a folder path rather than a file
pub(super) const FOLDER_PATTERNS: &[&str] = &["search-ms:", "shell:", "::{"];

/// Check if a name is a container/UI element name rather than an actual file
pub(super) fn is_container_name(name: &str) -> bool {
    if name.is_empty() {
        return true;
    }
    CONTAINER_NAMES
        .iter()
        .any(|&c| name.eq_ignore_ascii_case(c))
}

/// Check if a value looks like a valid file path (not a shell special path)
pub(super) fn is_valid_file_path(s: &str) -> bool {
    if s.is_empty() {
        return false;
    }
    // Skip shell special paths
    for pattern in FOLDER_PATTERNS {
        if s.to_lowercase().contains(pattern) {
            return false;
        }
    }
    // Check if it looks like a file path
    let path = PathBuf::from(s);
    path.is_absolute()
}

/// The item the pointer is over, as the view's accessibility provider reports it,
/// or nothing when the pointer is over no item at all.
///
/// `measure_content` asks the walk for the item's own text as well, which a preview
/// is placed from. A plain probe passes `false`: it answers which file the pointer is
/// on, and a preview that keeps off that file's item asks for the text when it is
/// about to be shown — see `avoid_box_under_cursor`.
pub(super) fn uia_item_from_point(
    resolver: &ItemResolver,
    point: POINT,
    measure_content: bool,
) -> Option<HoveredItem> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;
    note_probe(&PROBE_ITEM_WALKS);

    // Timed from the first crossing to the last — the element is asked of the view's
    // provider and the walk climbs from it — and timed around the early answers too,
    // which a `?` reaching out of the function would have timed past.
    let started = Instant::now();
    let found = (|| {
        let element = unsafe { automation.ElementFromPointBuildCache(point, cache) }.ok()?;
        walk_to_item(resolver, &element, Some(point), measure_content)
    })();
    note_probe_ms(&PROBE_ITEM_SLOWEST_MS, started.elapsed());

    found
}

/// The item the keyboard is on: the element Explorer says holds the focus, or the
/// item of that element's list when the view reports the list itself as focused.
pub(super) fn uia_item_from_focus(resolver: &ItemResolver) -> Option<HoveredItem> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;
    note_probe(&PROBE_ITEM_WALKS);

    let started = Instant::now();
    let found = (|| {
        let focused = unsafe { automation.GetFocusedElementBuildCache(cache) }.ok()?;

        // A view can report the list itself as focused while the focus it draws is on
        // one of its items — the search results view does — and a list's name is not a
        // file's, so the selection is where the item has to be taken from. For a
        // focused element that already is an item this is the element itself.
        let start = if element_names_an_item(&focused) {
            focused
        } else {
            selected_item_of_focused_list(&focused)?
        };

        // A keyboard preview is placed from the item alone, so the text the region is
        // made of is read with it — see `item_text_box`.
        walk_to_item(resolver, &start, None, true)
    })();
    note_probe_ms(&PROBE_ITEM_SLOWEST_MS, started.elapsed());

    found
}

/// The nearest item at or above an element, as the view reports it.
///
/// The walk starts where the caller's evidence does — the element under the
/// pointer, or the focused element — and goes up to the first list row or data
/// item, because what the view says about an item is kept on the item and not on
/// the text it draws inside it.
pub(super) fn walk_to_item(
    resolver: &ItemResolver,
    start: &IUIAutomationElement,
    point: Option<POINT>,
    measure_content: bool,
) -> Option<HoveredItem> {
    let cache = resolver.cache.as_ref()?;
    let walker = resolver.walker.as_ref()?;
    let mut element = start.clone();

    for _ in 0..=POINTER_ITEM_ANCESTOR_LIMIT {
        if let Some(mut item) = item_from_element(resolver, &element, point, measure_content) {
            // An item's box is the place its content is drawn, which for an item that has
            // been scrolled partly out of the view is not the place the view shows it: the
            // box carries on behind the toolbar above and past the edge of the window
            // below, and the pointer is on the search bar, the address bar, or off the
            // window there rather than on the item — a preview that nothing takes down,
            // because the box it is held by says the pointer never left. The view the item
            // is drawn in is the part of that box that is really there (see
            // `view_bounds`), so the box is kept to it here, once, and every question
            // asked of the item's box afterwards is asked of the part that exists.
            if let Some(view) = view_bounds(resolver, &element) {
                item.bounds = clip_box(item.bounds, view);
            }

            return Some(item);
        }

        element = unsafe { walker.GetParentElementBuildCache(&element, cache) }.ok()?;
    }

    None
}

/// The box of the view that shows an item: the element the item sits in, a level or two
/// up — the list itself, or the group a view that groups its items draws it in.
///
/// This is what clips an item's box to what is on screen. A provider reports an item's
/// box at the place the item's content is drawn however far the view has been scrolled,
/// so the box of an item half out of a view whose content is larger than its window
/// reaches behind the toolbar and past the bottom edge of the window. The container the
/// view draws its items in does not scroll with them: its rectangle is the window the
/// items are shown in, and an item is on screen exactly where the two overlap.
///
/// Nothing is answered when the walk cannot get there, which leaves the item's own box as
/// the provider reported it — the reading this app had before the container was asked.
pub(super) fn view_bounds(resolver: &ItemResolver, element: &IUIAutomationElement) -> Option<RECT> {
    let cache = resolver.cache.as_ref()?;
    let walker = resolver.walker.as_ref()?;
    let mut current = unsafe { walker.GetParentElementBuildCache(element, cache) }.ok()?;

    // A grouped view puts a group between the item and the list, and a group scrolls with
    // its items — so it is not the window they are shown in, and the climb goes on. The
    // climb is bounded like every other walk here: a provider that answers something
    // unexpected must not cost the probe an unbounded one.
    for _ in 0..POINTER_ITEM_ANCESTOR_LIMIT {
        match element_control_type(&current) {
            Some(control_type) if control_type == UIA_GroupControlTypeId => {
                current = unsafe { walker.GetParentElementBuildCache(&current, cache) }.ok()?;
            }
            _ => break,
        }
    }

    element_bounds(&current)
}

/// An item's box, kept to the part of it the view shows.
pub(super) fn clip_box(bounds: RECT, view: RECT) -> RECT {
    RECT {
        left: bounds.left.max(view.left),
        top: bounds.top.max(view.top),
        right: bounds.right.min(view.right),
        bottom: bounds.bottom.min(view.bottom),
    }
}

/// The item an element is, when the pointer's box test says it is the one asked
/// about — or, for the keyboard, whenever it is an item at all.
///
/// The box is what says the pointer is on an item: a row is an item from its left
/// edge to its right one, and a pointer anywhere on that row belongs to the file
/// the row stands for, which is the same thing Explorer highlights when it is
/// selected. An item whose box does not hold the point is not an answer even
/// though the walk passed through it, and since the walk starts at the element
/// under the pointer, the first item that holds the point is the one it is on. A
/// keyboard preview has no point to test — the focused element is the evidence —
/// so there it is the item itself that answers.
pub(super) fn item_from_element(
    resolver: &ItemResolver,
    element: &IUIAutomationElement,
    point: Option<POINT>,
    measure_content: bool,
) -> Option<HoveredItem> {
    let control_type = element_control_type(element)?;
    if control_type != UIA_ListItemControlTypeId && control_type != UIA_DataItemControlTypeId {
        return None;
    }

    let bounds = element_bounds(element)?;
    if bounds.left >= bounds.right || bounds.top >= bounds.bottom {
        return None;
    }
    if let Some(point) = point {
        let holds_point = point.x >= bounds.left
            && point.x < bounds.right
            && point.y >= bounds.top
            && point.y < bounds.bottom;
        if !holds_point {
            return None;
        }
    }

    // The item's own text is read only where it can answer, and only when the caller
    // wants it: what the view draws is what a way of avoiding is measured from. The
    // pointer's walk asks for it only while the setting keeps something off (see
    // `avoid_box_under_cursor`), where the keyboard's asks whenever it walks, because
    // the region a keyboard preview is placed from is read even at `Avoid Nothing` (see
    // `keyboard_avoid_box`). See `item_text_box`.
    let text = measure_content
        .then(|| item_text_box(resolver, element, &bounds))
        .flatten();

    Some(HoveredItem {
        index: element_item_index(element, resolver.item_index_property),
        name: element_name(element).unwrap_or_default().trim().to_string(),
        value: element_value(element),
        bounds,
        text,
        native_window: element_native_window(element),
    })
}

/// The text an item draws inside its own box, or `None` when it draws none.
///
/// A view gives every item the box it occupies, and what it draws inside that box
/// is reported the way a view reports everything: each piece of an item's text — a
/// Content row's name and path, a Details row's name, type, modified date and size,
/// the label under an icon — is an element of its own carrying the box it is drawn
/// in, child of the item. Reading them answers three things a box cannot: the whole of
/// what the item draws, whose right edge is where a row's content stops — the region
/// `Avoid Details` keeps a preview off, and the room *past* it a keyboard preview is
/// placed in — the leftmost piece, which in the views that draw their items as rows is
/// the `Name` column of `Details` or the name above the path of `Content`, and which
/// is the region a way of avoiding is measured from, and whether anything is drawn
/// beside that piece, which is what says the item is a row of its view at all — see
/// [`ItemText`].
///
/// It is measured from the item's own children for the same reason: a view reports
/// what it draws as children, and the rightmost of them is the edge the row's content
/// stops at. The read is batched into one round trip with the properties the rest of
/// the walk already asks for. A view that draws its columns some other way reports no
/// text at all, and is answered with `None`, which leaves the item measured by its box
/// — see [`HoveredItem::avoid_box`].
pub(super) fn item_text_box(
    resolver: &ItemResolver,
    element: &IUIAutomationElement,
    bounds: &RECT,
) -> Option<ItemText> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;

    let mut pieces: Vec<RECT> = Vec::new();

    unsafe {
        let condition = automation.CreateTrueCondition().ok()?;
        let children = element
            .FindAllBuildCache(TreeScope_Children, &condition, cache)
            .ok()?;
        let count = children.Length().ok()?;

        for index in 0..count {
            let Ok(child) = children.GetElement(index) else {
                continue;
            };
            if !element_is_drawn_text(&child) {
                continue;
            }
            let Ok(rect) = child.CachedBoundingRectangle() else {
                continue;
            };
            if rect.right <= rect.left {
                continue;
            }
            // Text reported outside the item's box is reported wrong, and the box is
            // what answers for an item whose text the view does not place.
            if rect.right > bounds.left && rect.left < bounds.right {
                pieces.push(rect);
            }
        }
    }

    text_boxes(&pieces)
}

/// What the pieces of an item's own text add up to, or `None` when the item drew none.
///
/// Three things are read from them at once, because one walk answers all three: the
/// whole of what the item draws as one box, the piece its name is drawn in, and
/// whether anything was drawn *beside* that piece. The last is what tells a row of the
/// view from a box item: a `Details` or `Content` row writes its columns to the right
/// of the name — the type, the date, the size — where a label under an icon, a tile's
/// stacked lines and a name on its own draw nothing there. See [`ItemText::columns`].
pub(super) fn text_boxes(pieces: &[RECT]) -> Option<ItemText> {
    let mut all: Option<RECT> = None;
    let mut name: Option<RECT> = None;

    for rect in pieces {
        // The name is the leftmost piece, which is the one the views that draw their
        // items as rows put first; two pieces drawn from the same edge — the name above
        // the path of `Content` — are told apart by taking the higher.
        let is_name = match name {
            None => true,
            Some(current) => {
                rect.left < current.left || (rect.left == current.left && rect.top < current.top)
            }
        };
        if is_name {
            name = Some(*rect);
        }

        all = Some(match all {
            Some(union) => RECT {
                left: union.left.min(rect.left),
                top: union.top.min(rect.top),
                right: union.right.max(rect.right),
                bottom: union.bottom.max(rect.bottom),
            },
            None => *rect,
        });
    }

    let all = all?;
    let name = name.unwrap_or(all);

    Some(ItemText {
        all,
        name,
        columns: all.right > name.right,
    })
}

/// Whether an element is text the view draws, which is what an item's own content
/// is made of. A row of a file list reports its name and its columns that way, and
/// anything else it may report — the file's icon, the row's own container — is not
/// part of the text whose end is being measured.
pub(super) fn element_is_drawn_text(element: &IUIAutomationElement) -> bool {
    match element_control_type(element) {
        Some(control_type) => {
            control_type == UIA_EditControlTypeId || control_type == UIA_TextControlTypeId
        }
        None => false,
    }
}

/// The control type of an element, from the batched cache when the element came
/// from one and from the provider itself when it did not: the item a keyboard
/// preview is resolved from can be handed over by a selection pattern, which is
/// not a cached read of the element under a pointer.
pub(super) fn element_control_type(element: &IUIAutomationElement) -> Option<UIA_CONTROLTYPE_ID> {
    unsafe {
        element
            .CachedControlType()
            .or_else(|_| element.CurrentControlType())
            .ok()
    }
}

pub(super) fn element_name(element: &IUIAutomationElement) -> Option<String> {
    unsafe {
        element
            .CachedName()
            .or_else(|_| element.CurrentName())
            .ok()
            .map(|name| name.to_string())
    }
}

pub(super) fn element_bounds(element: &IUIAutomationElement) -> Option<RECT> {
    unsafe {
        element
            .CachedBoundingRectangle()
            .or_else(|_| element.CurrentBoundingRectangle())
            .ok()
    }
}

pub(super) fn element_native_window(element: &IUIAutomationElement) -> isize {
    unsafe {
        element
            .CachedNativeWindowHandle()
            .or_else(|_| element.CurrentNativeWindowHandle())
            .map(|window| window.0 as isize)
            .unwrap_or_default()
    }
}

/// The value the item's legacy accessible pattern carries — for a file a search
/// has surfaced, the file's own path.
pub(super) fn element_value(element: &IUIAutomationElement) -> Option<String> {
    unsafe {
        let pattern = element
            .GetCachedPatternAs::<IUIAutomationLegacyIAccessiblePattern>(
                UIA_LegacyIAccessiblePatternId,
            )
            .or_else(|_| {
                element.GetCurrentPatternAs::<IUIAutomationLegacyIAccessiblePattern>(
                    UIA_LegacyIAccessiblePatternId,
                )
            })
            .ok()?;

        pattern
            .CachedValue()
            .or_else(|_| pattern.CurrentValue())
            .ok()
            .map(|value| value.to_string())
            .filter(|value| !value.trim().is_empty())
    }
}

/// The position the view holds an element at, as Explorer reports it.
///
/// The property is one-based, and zero is what the provider answers when it has
/// nothing to say about the element's position at all — so only a positive value
/// is a position. A missing one is not a failure: the item's own value still
/// stands on its own.
pub(super) fn element_item_index(
    element: &IUIAutomationElement,
    item_index_property: Option<UIA_PROPERTY_ID>,
) -> Option<i32> {
    let property = item_index_property?;

    unsafe {
        let value = element
            .GetCachedPropertyValue(property)
            .or_else(|_| element.GetCurrentPropertyValue(property))
            .ok()?;
        let raw = value.as_raw().Anonymous.Anonymous;
        if raw.vt != VT_I4.0 {
            return None;
        }

        let index = raw.Anonymous.lVal;
        (index > 0).then_some(index)
    }
}
