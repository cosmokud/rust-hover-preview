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
