# Rust Hover Preview

A Windows Explorer hover-preview utility where a tray host senses pointer and keyboard input while separate viewer processes render previews and standalone pinned windows.

## Language

### Processes & Windows

**Tray Host**:
The long-lived process that owns the system tray icon, Explorer input sensing, configuration, and update checks.
_Avoid_: Main app, controller, parent process

**Viewer**:
A rendering process that displays previews on behalf of the Tray Host.
_Avoid_: Renderer, worker, child window

**Serving Viewer**:
The single Viewer currently receiving hover previews and active pin commands from the Tray Host.
_Avoid_: Active viewer, primary viewer, hover process

**Pin**:
A preview frozen into a persistent interactive window that remains on screen after the pointer moves away from the file listing.
_Avoid_: Sticky preview, locked preview, modal preview

**Detached Window**:
A Pin that has been promoted into an independent window owned by its own Viewer process, unaffected by subsequent hovers or listing selections.
_Avoid_: Pop-out window, floating pin, hand-off window

### Cross-Process Coordination

**Seam**:
The shared coordination surface through which the Tray Host and its Viewers exchange commands, notices, and window state.
_Avoid_: IPC channel, bridge, shared buffer

**Viewer Slot**:
A per-Viewer record within the Seam holding one Viewer's published window geometry, input counters, and pin state.
_Avoid_: Block, bucket, lane

**Slot Reclaim**:
The transfer of a Viewer Slot previously held by an exited Viewer to a newly started Viewer after clearing all prior state.
_Avoid_: Slot reuse, slot takeover

**Viewer Generation**:
A monotonically increasing identity assigned to each Serving Viewer launch to distinguish the current Serving Viewer from superseded launch attempts.
_Avoid_: Viewer epoch, launch version, spawn id

**Graceful Shutdown**:
A coordinated exit in which the Tray Host notifies all active Viewers to release their external engines and media players before the session ends.
_Avoid_: Clean close, soft kill

**Idle Gate**:
The condition that no Pin, Detached Window, active hover hold, embedded document page, dialog, or menu is open, permitting a silent auto-update to proceed without interrupting the user.
_Avoid_: Quiet check, update lock

### Configuration & Propagation

**Configuration Publication**:
The single ordered act of updating settings in memory, saving them to disk, advancing the shared revision counter, and announcing the change to Viewers before any dependent command is sent.
_Avoid_: Config flush, settings sync, save-and-notify

**Configuration Generation**:
A monotonically increasing revision number published across the Seam whenever settings change, allowing both the Serving Viewer and Detached Windows to reload immediately rather than waiting for periodic disk polling.
_Avoid_: Config epoch, settings version, mtime token

**Configuration Snapshot**:
A single point-in-time copy of all settings taken once when constructing the tray menu so every row and submenu reflects a consistent state.
_Avoid_: Config clone, settings copy, menu cache

### Hover & Pin Sessions

**Hover Session**:
The active lifecycle of a single hovered file in a Viewer, owning the displayed file, its pending replay or failure state, and its resolved file assessment until the file changes, hides, or is promoted into a Pin.
_Avoid_: Preview flags, hover bookkeeping, replay slots

**Hover Snapshot**:
The point-in-time assessment of a hovered file's kind, drawing path, and active scale preferences carried from initial hover layout into Pin promotion so both views agree on how the file is measured and sized.
_Avoid_: Cached probe, file metadata copy, layout facts

### Explorer Sensing & Low-Latency Input

**Hook Session**:
The owned sensing state of the Explorer hook loop — active mouse and keyboard hovers, hovered item bounds, folder-change suspension gate, scroll tracking, and page-drag state — whose transitions clear dependent fields atomically so stale state cannot survive a reset.
_Avoid_: Loop locals, hook flags, polling state

**Single-Tick Pointer Reading**:
One consistent capture of cursor position, display scale, and window under the pointer taken once per tick and shared across both hover sensing and Pin follow-up logic.
_Avoid_: Double cursor read, fresh hit test, pointer re-poll

**Trigger Chord Latch**:
The keyboard hook's record of a trigger shortcut's main key combined with live verification that all companion modifier keys are also held, preventing a lone main key from falsely blocking or enabling previews.
_Avoid_: Sticky trigger, key bit, modifier cache

### Pin Surface & Film Presentation

**Pin Surface**:
The unified boundary through which a Pin's chrome, collapsed bubble, focus handover, and media installation are coordinated without coupling lifecycle decisions to desktop window handles.
_Avoid_: Pin handle wrapper, window helper, UI bridge

**Film Host**:
The dedicated presentation surface that owns a pinned video's chrome, carries its taskbar identity, and composites the live video feed cleanly across resizes, swaps, and display changes.
_Avoid_: Video container, player wrapper, backdrop window

**Pin Bubble**:
The compact circular desktop stand-in left behind when a Pin is collapsed, with an explicit lifecycle from creation through teardown so no orphaned desktop surface survives exit.
_Avoid_: Floating dot, mini pin, tray orb

### External Engines & Embedded Browser Host

**Engine Supervisor**:
The shared control layer for out-of-process converters that enforces hidden process launch, newest-wins work slots, give-up termination, strict budget-bounded output capture, and automatic staged-file cleanup.
_Avoid_: Process runner, CLI wrapper, converter helper

**Strict Budget Capture**:
A bounded read of converter output or staged files that fails explicitly when output exceeds the decode budget rather than silently returning truncated data.
_Avoid_: Capped read, truncated stream, soft limit

**WebView2 Host**:
The shared single-threaded browser hosting foundation — apartment lifecycle, message pump, host window creation, bounds synchronization, and navigation reveal state — built upon by both hover pages and pinned PDF viewers.
_Avoid_: Browser wrapper, WebView helper, Chromium shell

### Caching & Chrome Measurement

**Bounded Cache**:
A fixed-capacity in-memory memo that evicts the least-recently-used entry when full rather than discarding all remembered answers at once.
_Avoid_: Clearing map, cliff cache, ad-hoc memo

**Deterministic Disk Key**:
A stable digest derived from a file's identity and version so cached disk entries remain valid across application restarts and releases.
_Avoid_: Runtime hash, unstable key, session hash

**Display Scale**:
A sanitized representation of a monitor's scaling factor that falls back to standard scale when a display reports no valid density, preventing zero-sized chrome layout.
_Avoid_: Raw DPI ratio, unchecked scale, pixel divider

**Text Measurement Seam**:
The boundary through which chrome layout measures label widths and character runs independently of desktop drawing surfaces, allowing pooled measurement at runtime and deterministic headless testing.
_Avoid_: Scratch surface, throwaway DC, inline GDI measure
