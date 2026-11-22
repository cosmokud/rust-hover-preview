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
