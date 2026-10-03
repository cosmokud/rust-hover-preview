//! The shell's own objects, and the only part that holds one: the UI Automation client,
//! the window collection behind it, the view that answered last for a window, and the
//! memo of what a point was answered with.
//!
//! It is also what a walk is asked for a fresh one. These parts outlive the view being
//! recreated under them — a listing is rebuilt far more often than the automation client
//! behind it — so a walk that kept its own would be holding an object the shell has
//! already dropped, and the next walk would be reading through a dead one.

use super::*;

impl ItemResolver {
    pub(super) fn new(automation: Option<AutomationClient>) -> Self {
        let item_index_property = register_item_index_property();
        let automation_bounded = automation.as_ref().is_some_and(|client| client.bounded);

        let mut resolver = Self {
            automation: automation.map(|client| client.client),
            automation_bounded,
            cache: None,
            walker: None,
            item_index_property,
            shell_windows: None,
            window_views: None,
            item: None,
            probe: None,
        };
        resolver.rebuild_automation_parts();
        resolver.rebuild_shell();
        resolver
    }

    /// The batched property request and the tree walker, built from the client in
    /// hand. They are rebuilt with that client rather than kept across one, so
    /// nothing built by a client is read through its replacement.
    fn rebuild_automation_parts(&mut self) {
        let (cache, walker) = match self.automation.as_ref() {
            Some(automation) => unsafe {
                let cache = automation.CreateCacheRequest().ok();
                if let Some(cache) = cache.as_ref() {
                    let _ = cache.SetTreeScope(TreeScope_Element);
                    for property in [
                        UIA_ControlTypePropertyId,
                        UIA_BoundingRectanglePropertyId,
                        UIA_NativeWindowHandlePropertyId,
                        UIA_NamePropertyId,
                    ] {
                        let _ = cache.AddProperty(property);
                    }
                    if let Some(item_index) = self.item_index_property {
                        let _ = cache.AddProperty(item_index);
                    }
                    let _ = cache.AddPattern(UIA_LegacyIAccessiblePatternId);
                }
                (cache, automation.ControlViewWalker().ok())
            },
            None => (None, None),
        };

        self.cache = cache;
        self.walker = walker;
    }

    /// Ask for a client that can be bounded, for a resolver holding one that cannot.
    ///
    /// A client that could not be bounded is reached through the legacy object the
    /// timeouts are not on, so every probe made through it runs until the shell
    /// answers. Two states arrive here and only one of them is expected to leave.
    ///
    /// A resolver that never got a client at all is the one this recovers outright,
    /// and that is the work it is really for: a client that failed once would
    /// otherwise never be asked for again for the life of the run, and every hover
    /// after it would have no answer at all.
    ///
    /// A resolver holding an unbounded client is asked again on a cadence that
    /// grows, in case a bounded one can be made later. What that does *not* cover is
    /// worth being plain about, because it is the whole of what this can do: the
    /// modern object is served by the UI Automation core in this process rather than
    /// by the shell, so a machine that cannot create it is not a shell that is slow
    /// to come up — it is a machine without it, where the ask never succeeds and the
    /// unbounded client is kept. The timeouts cannot be added to a client that does
    /// not carry them, and nothing in user mode can interrupt a probe already inside
    /// one.
    ///
    /// A client no better than the one in hand is not taken. What this is for is an
    /// upgrade, and replacing one unbounded client with another of its kind would
    /// throw away the request and the walker built from it for nothing.
    pub(super) fn rebuild_automation(&mut self) {
        if self.automation_bounded {
            return;
        }

        let Some(client) = automation_client() else {
            return;
        };

        if !client.bounded && self.automation.is_some() {
            return;
        }

        self.automation = Some(client.client);
        self.automation_bounded = client.bounded;
        self.rebuild_automation_parts();
    }

    /// Build the Shell window collection again, and drop what was read through it: the
    /// views of the window last resolved in, and the item under the pointer. These are
    /// what the resolver holds that another process serves: the collection and every view
    /// are Explorer's own, so once Explorer is not the process it was, both are proxies
    /// into one that is gone — and a proxy into a gone process fails for good rather than
    /// reconnecting.
    pub(super) fn rebuild_shell(&mut self) {
        self.shell_windows =
            unsafe { CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL).ok() };
        self.window_views = None;
        self.item = None;
    }

    /// Drop a window's views, for a caller that has just seen the place they describe
    /// change. A set is otherwise kept across a navigation — a view is the same object
    /// afterwards and reads the folder it holds now — so this is for the change nothing
    /// about the set can be read against, and what it costs is one walk.
    pub(super) fn forget_window_views(&mut self) {
        self.window_views = None;
    }

    /// The view of `frame` the pointer was last inside, out of the set kept for that frame.
    ///
    /// What reads it is a probe that is in none of the frame's views and has to answer with
    /// something anyway — a pointer over the navigation pane or the toolbar — and the
    /// answer is the tab the hand was last working in rather than one picked at random (see
    /// `WindowViews`).
    pub(super) fn remembered_view(&self, frame: isize) -> Option<usize> {
        self.window_views
            .as_ref()
            .filter(|set| set.frame == frame)
            .and_then(|set| set.anchor)
    }

    /// Remember that a probe found the pointer inside this view of this frame.
    pub(super) fn remember_view(&mut self, frame: isize, index: usize) {
        if let Some(set) = self.window_views.as_mut() {
            if set.frame == frame {
                set.anchor = Some(index);
            }
        }
    }

    /// Drop the item under the pointer, for everything that makes it a new question: a
    /// wheel moving the list under a parked pointer, a navigation, a click that may have
    /// sorted the view. What is kept is an answer for the item it was read from, and none
    /// of those leave that item where it was.
    pub(super) fn forget_item(&mut self) {
        self.item = None;
    }

    /// One answer per loop tick: what a tick learned is not carried into the next
    /// one, where the list under a parked pointer may have moved on.
    pub(super) fn forget_probe(&mut self) {
        self.probe = None;
    }

    pub(super) fn probed_at(&self, point: POINT) -> Option<Option<PathBuf>> {
        self.probe
            .as_ref()
            .filter(|probe| probe.point.x == point.x && probe.point.y == point.y)
            .map(|probe| probe.answer.clone())
    }

    pub(super) fn remember_probe(&mut self, point: POINT, answer: Option<PathBuf>) {
        self.probe = Some(ProbeMemo { point, answer });
    }

    /// The answer the item under the pointer was read for, where the pointer is still
    /// inside that item: the same window, and the same box within it — see `AnsweredItem`.
    pub(super) fn item_under(&self, point: POINT, drawn_in: isize) -> Option<Option<PathBuf>> {
        self.item
            .as_ref()
            .filter(|answered| {
                answered.drawn_in == drawn_in && point_in_box(point, answered.bounds)
            })
            .map(|answered| answered.path.clone())
    }

    /// Keep what one look found, against the item it read rather than the point it read
    /// it at. A look that found no item at all keeps nothing: what a look like that
    /// answered for is the asking having failed, and that is a question to ask again next
    /// tick (see `read_failure_is_the_same_item`).
    pub(super) fn remember_item(&mut self, look: &PointerLook) {
        self.item = look.item_bounds.map(|bounds| AnsweredItem {
            drawn_in: look.drawn_in,
            bounds,
            path: look.path.clone(),
        });
    }
}

/// How long a UI Automation call may take before it is abandoned. Every probe
/// crosses into the shell's own thread, so a shell that has stopped answering
/// would otherwise hold this thread inside a probe for as long as it likes — and
/// the slow-probe backoff cannot see a probe that has not returned yet, which
/// leaves an unbounded wait with nothing watching it.
pub(super) const UIA_TIMEOUT_MS: u32 = 500;

/// How long a resolver left without a client it could bound waits before it asks
/// for one again, and the ceiling that wait grows to; see
/// `ItemResolver::rebuild_automation`.
pub(super) const UIA_REBIND_RETRY_MS: u64 = 2000;
pub(super) const UIA_REBIND_INTERVAL_MAX_MS: u64 = 60_000;

/// The UI Automation client both paths resolve items with, and whether every call
/// it makes is bounded.
pub(super) struct AutomationClient {
    client: IUIAutomation,
    /// Whether `IUIAutomation2` answered for this client *and* took both timeouts.
    /// What a client that is not bounded costs is a probe that runs until the shell
    /// answers, however long that is — which is why it is a state to leave rather
    /// than one to settle into, and why it is carried beside the client rather than
    /// assumed from it.
    bounded: bool,
}

/// The UI Automation client both paths resolve items with, with every call
/// bounded wherever the machine can bound one.
///
/// `CUIAutomation8` is the client that carries `IUIAutomation2`, which is where
/// the timeouts live: the legacy `CUIAutomation` object does not answer for that
/// interface at all, so asking it to be bounded is what failed — and, when that
/// was treated as fatal, what left every hover without a preview. The legacy
/// client is still the fallback, and a client that cannot be bounded is still
/// used, because a client that cannot be bounded still resolves items: what it
/// costs is the wait the timeouts were meant to remove, which is the behavior the
/// app had before, and not a dead app.
///
/// What it is not is a client to keep. The fallback is reached exactly where
/// `CUIAutomation8` could not be created, and a shell that is not up yet is as good
/// a reason for that as a machine without the object is, so a client that came back
/// unbounded is asked for again rather than held — see
/// `ItemResolver::rebuild_automation`.
///
/// Both setters are checked rather than assumed. A client that answered for the
/// interface and refused a call would otherwise be recorded as bounded while every
/// probe it makes ran unbounded, which is the one thing this pair is here to stop.
pub(super) fn automation_client() -> Option<AutomationClient> {
    let client: IUIAutomation = unsafe {
        match CoCreateInstance(&CUIAutomation8, None, CLSCTX_ALL) {
            Ok(client) => client,
            Err(_) => CoCreateInstance(&CUIAutomation, None, CLSCTX_ALL).ok()?,
        }
    };

    let bounded = match client.cast::<IUIAutomation2>() {
        Ok(bounded) => unsafe {
            bounded.SetConnectionTimeout(UIA_TIMEOUT_MS).is_ok()
                && bounded.SetTransactionTimeout(UIA_TIMEOUT_MS).is_ok()
        },
        Err(_) => false,
    };

    Some(AutomationClient { client, bounded })
}

/// Explorer's own `ItemIndex` property, asked of the UI Automation registrar.
///
/// The id a registered property is read under is a runtime value, so the GUID
/// Explorer publishes under the name `ItemIndex` has to be exchanged for it once
/// before the property can be read at all. Everything about this is optional: a
/// registrar that refuses the property leaves the poke points above unread, and
/// the pointer is answered by the routes that do not need a position.
pub(super) fn register_item_index_property() -> Option<UIA_PROPERTY_ID> {
    unsafe {
        let registrar: IUIAutomationRegistrar =
            CoCreateInstance(&CUIAutomationRegistrar, None, CLSCTX_INPROC_SERVER).ok()?;
        let property = registrar
            .RegisterProperty(&UIAutomationPropertyInfo {
                guid: ItemIndex_Property_GUID,
                pProgrammaticName: w!("ItemIndex"),
                r#type: UIAutomationType_Int,
            })
            .ok()?;

        (property > 0).then_some(UIA_PROPERTY_ID(property))
    }
}

/// "Do not preview this file" latch, shared by the mouse hover path. It is a
/// delay, not a verdict: the file it names is held off the mouse path until the
/// same-file rehover delay has passed, and released sooner when the cursor
/// resolves another file. A keyboard preview that a mouse move dismissed latches
/// the file it showed the same way — long enough that the handover cannot flash
/// the file straight back, and no longer, because a pointer parked on that file
/// afterwards is a user asking for it and not a repeat of the handover.
#[derive(Default)]
pub(super) struct SuppressedHover {
    pub(super) file: Option<PathBuf>,
    started_at: Option<Instant>,
}

impl SuppressedHover {
    pub(super) fn clear(&mut self) {
        self.file = None;
        self.started_at = None;
    }

    pub(super) fn suppress(&mut self, file: PathBuf) {
        self.file = Some(file);
        self.started_at = Some(Instant::now());
    }

    pub(super) fn matches(&self, path: &Path) -> bool {
        self.file
            .as_ref()
            .map(|file| same_path(file, path))
            .unwrap_or(false)
    }

    pub(super) fn rehover_allowed(&self, required_delay_ms: u64) -> bool {
        self.started_at
            .map(|started| started.elapsed() >= Duration::from_millis(required_delay_ms))
            .unwrap_or(true)
    }
}

/// Pointer freeze for keyboard previews. A keyboard preview is placed next to
/// the focused item, which can put it right over the parked cursor; the pointer
/// must not take over in that case. The freeze is decided from the preview's own
/// box at spawn time and released by a real mouse move, so cursor jitter cannot
/// end a keyboard preview.
#[derive(Default)]
pub(super) struct KeyboardPointerPause {
    armed: bool,
    /// Set on every keyboard preview spawn while the preview thread has not
    /// published a box yet.
    box_watch_until: Option<Instant>,
    /// Last observed box, so a box that is still being replaced is not trusted.
    last_box: Option<(i32, i32, i32, i32)>,
}

impl KeyboardPointerPause {
    /// Called on every keyboard preview spawn: the box arrives once the preview
    /// is actually on screen, and the freeze is decided from it.
    pub(super) fn watch_for_box(&mut self) {
        self.armed = false;
        self.last_box = None;
        self.box_watch_until =
            Some(Instant::now() + Duration::from_millis(KEYBOARD_PREVIEW_BOX_WATCH_MS));
    }

    pub(super) fn clear(&mut self) {
        self.armed = false;
        self.box_watch_until = None;
        self.last_box = None;
    }

    pub(super) fn is_watching(&self) -> bool {
        self.box_watch_until.is_some()
    }

    /// True while the pointer must not drive previews, probe them, or dismiss
    /// them: the preview box is either known to cover the cursor, or still being
    /// waited on.
    pub(super) fn freezes_pointer(&self) -> bool {
        self.armed || self.box_watch_until.is_some()
    }

    /// Cursor movement that hands control back to the mouse. Small movements are
    /// ignored on purpose so a parked mouse cannot cancel a keyboard preview — and
    /// the wider tolerance holds for the whole of the keyboard's turn, not only
    /// while the preview's box happens to cover the cursor.
    pub(super) fn move_threshold_px(&self, keyboard_owns_screen: bool, dpi: u32) -> i32 {
        let logical_pixels = if keyboard_owns_screen || self.freezes_pointer() {
            KEYBOARD_POINTER_MOVE_TOLERANCE_PIXELS
        } else {
            MOUSE_MOVE_PIXELS
        };

        (logical_pixels * dpi as f32 / 96.0).round() as i32
    }

    /// Decide the freeze from the preview's on-screen box. A box only counts
    /// once it survives a second observation, because a preview that is being
    /// replaced can still report the outgoing window. `None` keeps waiting until
    /// the watch expires, so a preview that never appears cannot freeze the
    /// pointer forever.
    pub(super) fn evaluate_box(
        &mut self,
        cursor: POINT,
        preview_box: Option<(i32, i32, i32, i32)>,
    ) {
        let Some(watch_until) = self.box_watch_until else {
            return;
        };

        let Some(box_rect) = preview_box else {
            self.last_box = None;
            if Instant::now() >= watch_until {
                self.box_watch_until = None;
            }
            return;
        };

        if self.last_box != Some(box_rect) {
            self.last_box = Some(box_rect);
            return;
        }

        let (left, top, right, bottom) = box_rect;
        self.box_watch_until = None;
        self.last_box = None;
        self.armed = cursor.x >= left && cursor.x < right && cursor.y >= top && cursor.y < bottom;
    }
}
